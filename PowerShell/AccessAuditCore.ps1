Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-RequiredPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "No existe el archivo: $Path"
    }

    return (Resolve-Path -LiteralPath $Path).Path
}

function New-AccessConnection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath
    )

    $connectionString = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=$DatabasePath;Persist Security Info=False;"
    $connection = [System.Data.OleDb.OleDbConnection]::new($connectionString)
    $connection.Open()
    return $connection
}

function Get-UserTables {
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.OleDb.OleDbConnection]$Connection
    )

    $tables = $Connection.GetSchema("Tables")
    return $tables |
        Where-Object {
            $_.TABLE_TYPE -eq "TABLE" -and
            $_.TABLE_NAME -notlike "MSys*" -and
            $_.TABLE_NAME -notlike "~*"
        } |
        Sort-Object TABLE_NAME |
        Select-Object -ExpandProperty TABLE_NAME
}

function Get-AccessTableNames {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath
    )

    $resolvedPath = Resolve-RequiredPath -Path $DatabasePath
    $connection = $null

    try {
        $connection = New-AccessConnection -DatabasePath $resolvedPath
        return @(Get-UserTables -Connection $connection)
    }
    finally {
        if ($connection) {
            $connection.Dispose()
        }
    }
}

function Get-ColumnNames {
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.OleDb.OleDbConnection]$Connection,

        [Parameter(Mandatory = $true)]
        [string]$TableName
    )

    $command = $Connection.CreateCommand()
    $command.CommandText = "SELECT * FROM [$TableName] WHERE 1 = 0"
    $reader = $command.ExecuteReader()

    try {
        $schema = $reader.GetSchemaTable()
        return $schema |
            Sort-Object ColumnOrdinal |
            Select-Object -ExpandProperty ColumnName
    }
    finally {
        $reader.Close()
    }
}

function Get-PrimaryKeyColumns {
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.OleDb.OleDbConnection]$Connection,

        [Parameter(Mandatory = $true)]
        [string]$TableName
    )

    try {
        $catalog = New-Object -ComObject ADOX.Catalog
        $catalog.ActiveConnection = $Connection.ConnectionString
        $table = $catalog.Tables.Item($TableName)
        $primaryKey = $table.Keys | Where-Object { $_.Type -eq 1 } | Select-Object -First 1
        if (-not $primaryKey) {
            return @()
        }

        $columns = foreach ($column in $primaryKey.Columns) {
            $column.Name
        }

        return @($columns)
    }
    catch {
        return @()
    }
}

function Convert-FieldValue {
    param(
        $Value
    )

    if ($null -eq $Value -or $Value -is [System.DBNull]) {
        return $null
    }

    if ($Value -is [datetime]) {
        return $Value.ToString("o")
    }

    if ($Value -is [byte[]]) {
        return [Convert]::ToBase64String($Value)
    }

    return [string]$Value
}

function Get-RowMap {
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.OleDb.OleDbConnection]$Connection,

        [Parameter(Mandatory = $true)]
        [string]$TableName,

        [Parameter(Mandatory = $true)]
        [string[]]$Columns,

        [Parameter(Mandatory = $true)]
        [string[]]$KeyColumns
    )

    $command = $Connection.CreateCommand()
    $command.CommandText = "SELECT * FROM [$TableName]"
    $reader = $command.ExecuteReader()

    $rows = @{}
    $rowNumber = 0

    try {
        while ($reader.Read()) {
            $rowNumber++
            $row = [ordered]@{}
            foreach ($column in $Columns) {
                $row[$column] = Convert-FieldValue $reader[$column]
            }

            if ($KeyColumns.Count -gt 0) {
                $keyValues = foreach ($keyColumn in $KeyColumns) {
                    "$keyColumn=$($row[$keyColumn])"
                }
                $rowKey = $keyValues -join "|"
            }
            else {
                $fingerprintValues = foreach ($column in $Columns) {
                    "$column=$($row[$column])"
                }
                $rowKey = "ROW:${rowNumber}:" + ($fingerprintValues -join "|")
            }

            $rows[$rowKey] = $row
        }
    }
    finally {
        $reader.Close()
    }

    return $rows
}

function Compare-TableContent {
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.OleDb.OleDbConnection]$BeforeConnection,

        [Parameter(Mandatory = $true)]
        [System.Data.OleDb.OleDbConnection]$AfterConnection,

        [Parameter(Mandatory = $true)]
        [string]$TableName,

        [Parameter(Mandatory = $true)]
        [string[]]$IgnoreColumns
    )

    $beforeColumns = Get-ColumnNames -Connection $BeforeConnection -TableName $TableName
    $afterColumns = Get-ColumnNames -Connection $AfterConnection -TableName $TableName
    $commonColumns = @($beforeColumns | Where-Object { $afterColumns -contains $_ })
    $trackedColumns = @($commonColumns | Where-Object { $IgnoreColumns -notcontains $_ })
    $primaryKeyColumns = @(Get-PrimaryKeyColumns -Connection $AfterConnection -TableName $TableName)

    $beforeRows = Get-RowMap -Connection $BeforeConnection -TableName $TableName -Columns $trackedColumns -KeyColumns $primaryKeyColumns
    $afterRows = Get-RowMap -Connection $AfterConnection -TableName $TableName -Columns $trackedColumns -KeyColumns $primaryKeyColumns

    $allKeys = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($key in $beforeRows.Keys) { [void]$allKeys.Add($key) }
    foreach ($key in $afterRows.Keys) { [void]$allKeys.Add($key) }

    $inserted = [System.Collections.Generic.List[object]]::new()
    $deleted = [System.Collections.Generic.List[object]]::new()
    $modified = [System.Collections.Generic.List[object]]::new()

    foreach ($key in ($allKeys | Sort-Object)) {
        $hasBefore = $beforeRows.ContainsKey($key)
        $hasAfter = $afterRows.ContainsKey($key)

        if ($hasBefore -and -not $hasAfter) {
            $deleted.Add([pscustomobject]@{
                key = $key
                row = $beforeRows[$key]
            })
            continue
        }

        if (-not $hasBefore -and $hasAfter) {
            $inserted.Add([pscustomobject]@{
                key = $key
                row = $afterRows[$key]
            })
            continue
        }

        $beforeRow = $beforeRows[$key]
        $afterRow = $afterRows[$key]
        $fieldChanges = [System.Collections.Generic.List[object]]::new()

        foreach ($column in $trackedColumns) {
            $beforeValue = $beforeRow[$column]
            $afterValue = $afterRow[$column]
            if ($beforeValue -cne $afterValue) {
                $fieldChanges.Add([pscustomobject]@{
                    column = $column
                    before = $beforeValue
                    after = $afterValue
                })
            }
        }

        if ($fieldChanges.Count -gt 0) {
            $modified.Add([pscustomobject]@{
                key = $key
                changes = $fieldChanges
            })
        }
    }

    return [pscustomobject]@{
        table = $TableName
        primary_key_columns = $primaryKeyColumns
        tracked_columns = $trackedColumns
        has_reliable_key = ($primaryKeyColumns.Count -gt 0)
        inserted_count = $inserted.Count
        deleted_count = $deleted.Count
        modified_count = $modified.Count
        inserted = $inserted
        deleted = $deleted
        modified = $modified
    }
}

function Write-ConsoleSummary {
    param(
        [object[]]$TableReports
    )

    if ($null -eq $TableReports) {
        $TableReports = @()
    }

    $changedTables = @($TableReports | Where-Object {
        $_.inserted_count -gt 0 -or $_.deleted_count -gt 0 -or $_.modified_count -gt 0
    })

    if ($changedTables.Count -eq 0) {
        Write-Host "No se detectaron cambios en las tablas analizadas."
        return
    }

    Write-Host ""
    Write-Host "Tablas con cambios:"
    foreach ($table in $changedTables) {
        $keyMode = if ($table.has_reliable_key) { "PK" } else { "sin PK fiable" }
        Write-Host ("- {0}: +{1} / -{2} / ~{3} ({4})" -f $table.table, $table.inserted_count, $table.deleted_count, $table.modified_count, $keyMode)
    }
}

function Invoke-AccessAuditComparison {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BeforePath,

        [Parameter(Mandatory = $true)]
        [string]$AfterPath,

        $IgnoreColumns = @(
            "FechaModificacion",
            "Fecha_Modificacion",
            "FecMod",
            "UsuarioModificacion",
            "Usuario_Modificacion"
        ),

        [switch]$IncludeUnchangedTables,

        $TableNames
    )

    if ($null -eq $IgnoreColumns) {
        $IgnoreColumns = @()
    }
    else {
        $IgnoreColumns = @($IgnoreColumns | ForEach-Object { [string]$_ })
    }

    if ($null -eq $TableNames) {
        $TableNames = @()
    }
    else {
        $TableNames = @($TableNames | ForEach-Object { [string]$_ })
    }

    $resolvedBeforePath = Resolve-RequiredPath -Path $BeforePath
    $resolvedAfterPath = Resolve-RequiredPath -Path $AfterPath

    $beforeConnection = $null
    $afterConnection = $null

    try {
        $beforeConnection = New-AccessConnection -DatabasePath $resolvedBeforePath
        $afterConnection = New-AccessConnection -DatabasePath $resolvedAfterPath

        $beforeTables = @(Get-UserTables -Connection $beforeConnection)
        $afterTables = @(Get-UserTables -Connection $afterConnection)
        $allTables = @($beforeTables + $afterTables | Sort-Object -Unique)

        if ($TableNames -and @($TableNames).Count -gt 0) {
            $selectedTables = @($allTables | Where-Object { $TableNames -contains $_ })
        }
        else {
            $selectedTables = $allTables
        }

        $tableReports = [System.Collections.Generic.List[object]]::new()
        foreach ($tableName in $selectedTables) {
            if (($beforeTables -contains $tableName) -and ($afterTables -contains $tableName)) {
                $report = Compare-TableContent -BeforeConnection $beforeConnection -AfterConnection $afterConnection -TableName $tableName -IgnoreColumns $IgnoreColumns
            }
            elseif ($afterTables -contains $tableName) {
                $report = [pscustomobject]@{
                    table = $tableName
                    primary_key_columns = @()
                    tracked_columns = @()
                    has_reliable_key = $false
                    inserted_count = -1
                    deleted_count = 0
                    modified_count = 0
                    inserted = @()
                    deleted = @()
                    modified = @()
                    note = "La tabla solo existe en la base de datos DESPUES."
                }
            }
            else {
                $report = [pscustomobject]@{
                    table = $tableName
                    primary_key_columns = @()
                    tracked_columns = @()
                    has_reliable_key = $false
                    inserted_count = 0
                    deleted_count = -1
                    modified_count = 0
                    inserted = @()
                    deleted = @()
                    modified = @()
                    note = "La tabla solo existe en la base de datos ANTES."
                }
            }

            $hasChanges =
                $report.PSObject.Properties.Name -contains "note" -or
                $report.inserted_count -gt 0 -or
                $report.deleted_count -gt 0 -or
                $report.modified_count -gt 0

            if ($IncludeUnchangedTables -or $hasChanges) {
                $tableReports.Add($report)
            }
        }

        return [pscustomobject]@{
            generated_at = (Get-Date).ToString("o")
            before_database = $resolvedBeforePath
            after_database = $resolvedAfterPath
            ignored_columns = $IgnoreColumns
            table_count = $tableReports.Count
            tables = $tableReports
        }
    }
    finally {
        if ($beforeConnection) { $beforeConnection.Dispose() }
        if ($afterConnection) { $afterConnection.Dispose() }
    }
}
