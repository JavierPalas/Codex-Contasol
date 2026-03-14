Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. "$PSScriptRoot\\AccessAuditCore.ps1"

function New-Label {
    param(
        [string]$Text,
        [int]$X,
        [int]$Y,
        [int]$Width = 120
    )

    $label = [System.Windows.Forms.Label]::new()
    $label.Text = $Text
    $label.Location = [System.Drawing.Point]::new($X, $Y)
    $label.Size = [System.Drawing.Size]::new($Width, 24)
    return $label
}

function New-TextBox {
    param(
        [int]$X,
        [int]$Y,
        [int]$Width = 620
    )

    $textBox = [System.Windows.Forms.TextBox]::new()
    $textBox.Location = [System.Drawing.Point]::new($X, $Y)
    $textBox.Size = [System.Drawing.Size]::new($Width, 28)
    return $textBox
}

function New-Button {
    param(
        [string]$Text,
        [int]$X,
        [int]$Y,
        [int]$Width = 130
    )

    $button = [System.Windows.Forms.Button]::new()
    $button.Text = $Text
    $button.Location = [System.Drawing.Point]::new($X, $Y)
    $button.Size = [System.Drawing.Size]::new($Width, 32)
    return $button
}

function Show-ErrorMessage {
    param(
        [string]$Message
    )

    [System.Windows.Forms.MessageBox]::Show($Message, "Auditoria Access", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
}

function Show-InfoMessage {
    param(
        [string]$Message
    )

    [System.Windows.Forms.MessageBox]::Show($Message, "Auditoria Access", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
}

function Choose-AccessFile {
    $dialog = [System.Windows.Forms.OpenFileDialog]::new()
    $dialog.Filter = "Bases de datos Access (*.accdb;*.mdb)|*.accdb;*.mdb|Todos los archivos (*.*)|*.*"
    $dialog.Multiselect = $false

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }

    return $null
}

function Populate-TableCombo {
    param(
        [System.Windows.Forms.ComboBox]$Combo,
        [string[]]$Tables
    )

    $Combo.Items.Clear()
    [void]$Combo.Items.Add("(Todas)")
    foreach ($table in $Tables) {
        [void]$Combo.Items.Add($table)
    }
    $Combo.SelectedIndex = 0
}

function Format-RowDetails {
    param(
        [object]$Entry,
        [string]$Mode
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("Clave: $($Entry.key)")
    $lines.Add("")

    if ($Mode -eq "modified") {
        foreach ($change in $Entry.changes) {
            $lines.Add("Campo: $($change.column)")
            $lines.Add("  Antes: $($change.before)")
            $lines.Add("  Despues: $($change.after)")
            $lines.Add("")
        }
    }
    else {
        foreach ($property in $Entry.row.PSObject.Properties) {
            $lines.Add("$($property.Name): $($property.Value)")
        }
    }

    return ($lines -join [Environment]::NewLine).TrimEnd()
}

$form = [System.Windows.Forms.Form]::new()
$form.Text = "Auditoria de Access"
$form.StartPosition = "CenterScreen"
$form.Size = [System.Drawing.Size]::new(1180, 760)
$form.MinimumSize = [System.Drawing.Size]::new(1180, 760)

$font = [System.Drawing.Font]::new("Segoe UI", 9)
$form.Font = $font

$labelBefore = New-Label -Text "Base ANTES" -X 20 -Y 20
$txtBefore = New-TextBox -X 20 -Y 45
$btnBefore = New-Button -Text "Seleccionar..." -X 650 -Y 43

$labelAfter = New-Label -Text "Base DESPUES" -X 20 -Y 85
$txtAfter = New-TextBox -X 20 -Y 110
$btnAfter = New-Button -Text "Seleccionar..." -X 650 -Y 108

$labelBeforeTable = New-Label -Text "Tabla ANTES" -X 20 -Y 150
$cmbBeforeTable = [System.Windows.Forms.ComboBox]::new()
$cmbBeforeTable.Location = [System.Drawing.Point]::new(20, 175)
$cmbBeforeTable.Size = [System.Drawing.Size]::new(300, 28)
$cmbBeforeTable.DropDownStyle = "DropDownList"

$labelAfterTable = New-Label -Text "Tabla DESPUES" -X 340 -Y 150
$cmbAfterTable = [System.Windows.Forms.ComboBox]::new()
$cmbAfterTable.Location = [System.Drawing.Point]::new(340, 175)
$cmbAfterTable.Size = [System.Drawing.Size]::new(300, 28)
$cmbAfterTable.DropDownStyle = "DropDownList"

$labelIgnore = New-Label -Text "Ignorar columnas" -X 20 -Y 215 -Width 140
$txtIgnore = New-TextBox -X 20 -Y 240 -Width 620
$txtIgnore.Text = "FechaModificacion,Fecha_Modificacion,FecMod,UsuarioModificacion,Usuario_Modificacion"

$chkUnchanged = [System.Windows.Forms.CheckBox]::new()
$chkUnchanged.Text = "Incluir tablas sin cambios"
$chkUnchanged.Location = [System.Drawing.Point]::new(20, 278)
$chkUnchanged.Size = [System.Drawing.Size]::new(220, 24)

$btnLoadTables = New-Button -Text "Cargar tablas" -X 790 -Y 60 -Width 150
$btnAudit = New-Button -Text "Ejecutar auditoria" -X 950 -Y 60 -Width 170
$btnExport = New-Button -Text "Exportar JSON" -X 950 -Y 102 -Width 170
$btnExport.Enabled = $false

$lblSummary = [System.Windows.Forms.Label]::new()
$lblSummary.Text = "Resumen de cambios"
$lblSummary.Location = [System.Drawing.Point]::new(20, 320)
$lblSummary.Size = [System.Drawing.Size]::new(220, 24)

$gridTables = [System.Windows.Forms.DataGridView]::new()
$gridTables.Location = [System.Drawing.Point]::new(20, 350)
$gridTables.Size = [System.Drawing.Size]::new(540, 340)
$gridTables.ReadOnly = $true
$gridTables.AllowUserToAddRows = $false
$gridTables.AllowUserToDeleteRows = $false
$gridTables.SelectionMode = "FullRowSelect"
$gridTables.MultiSelect = $false
$gridTables.AutoSizeColumnsMode = "Fill"
$gridTables.RowHeadersVisible = $false

$lblDetails = [System.Windows.Forms.Label]::new()
$lblDetails.Text = "Detalle de la tabla seleccionada"
$lblDetails.Location = [System.Drawing.Point]::new(580, 320)
$lblDetails.Size = [System.Drawing.Size]::new(250, 24)

$tabsDetails = [System.Windows.Forms.TabControl]::new()
$tabsDetails.Location = [System.Drawing.Point]::new(580, 350)
$tabsDetails.Size = [System.Drawing.Size]::new(560, 340)

$tabInserted = [System.Windows.Forms.TabPage]::new()
$tabInserted.Text = "Altas"
$tabDeleted = [System.Windows.Forms.TabPage]::new()
$tabDeleted.Text = "Bajas"
$tabModified = [System.Windows.Forms.TabPage]::new()
$tabModified.Text = "Modificados"

$listInserted = [System.Windows.Forms.ListBox]::new()
$listInserted.Dock = "Left"
$listInserted.Width = 200
$txtInserted = [System.Windows.Forms.TextBox]::new()
$txtInserted.Dock = "Fill"
$txtInserted.Multiline = $true
$txtInserted.ScrollBars = "Vertical"
$txtInserted.ReadOnly = $true
$tabInserted.Controls.Add($txtInserted)
$tabInserted.Controls.Add($listInserted)

$listDeleted = [System.Windows.Forms.ListBox]::new()
$listDeleted.Dock = "Left"
$listDeleted.Width = 200
$txtDeleted = [System.Windows.Forms.TextBox]::new()
$txtDeleted.Dock = "Fill"
$txtDeleted.Multiline = $true
$txtDeleted.ScrollBars = "Vertical"
$txtDeleted.ReadOnly = $true
$tabDeleted.Controls.Add($txtDeleted)
$tabDeleted.Controls.Add($listDeleted)

$listModified = [System.Windows.Forms.ListBox]::new()
$listModified.Dock = "Left"
$listModified.Width = 200
$txtModified = [System.Windows.Forms.TextBox]::new()
$txtModified.Dock = "Fill"
$txtModified.Multiline = $true
$txtModified.ScrollBars = "Vertical"
$txtModified.ReadOnly = $true
$tabModified.Controls.Add($txtModified)
$tabModified.Controls.Add($listModified)

$tabsDetails.TabPages.AddRange(@($tabInserted, $tabDeleted, $tabModified))

$status = [System.Windows.Forms.StatusStrip]::new()
$statusLabel = [System.Windows.Forms.ToolStripStatusLabel]::new()
$statusLabel.Text = "Selecciona las bases de datos y pulsa 'Cargar tablas'."
$status.Items.Add($statusLabel) | Out-Null

$form.Controls.AddRange(@(
    $labelBefore, $txtBefore, $btnBefore,
    $labelAfter, $txtAfter, $btnAfter,
    $labelBeforeTable, $cmbBeforeTable,
    $labelAfterTable, $cmbAfterTable,
    $labelIgnore, $txtIgnore, $chkUnchanged,
    $btnLoadTables, $btnAudit, $btnExport,
    $lblSummary, $gridTables, $lblDetails, $tabsDetails, $status
))

Populate-TableCombo -Combo $cmbBeforeTable -Tables @()
Populate-TableCombo -Combo $cmbAfterTable -Tables @()

$currentResult = $null
$currentTableReports = @()
$currentInserted = @()
$currentDeleted = @()
$currentModified = @()

function Load-TableLists {
    if ([string]::IsNullOrWhiteSpace($txtBefore.Text) -or [string]::IsNullOrWhiteSpace($txtAfter.Text)) {
        throw "Debes seleccionar la base ANTES y la base DESPUES."
    }

    $beforeTables = @(Get-AccessTableNames -DatabasePath $txtBefore.Text)
    $afterTables = @(Get-AccessTableNames -DatabasePath $txtAfter.Text)
    Populate-TableCombo -Combo $cmbBeforeTable -Tables $beforeTables
    Populate-TableCombo -Combo $cmbAfterTable -Tables $afterTables
    $statusLabel.Text = "Tablas cargadas. Ya puedes ejecutar la auditoria."
}

function Get-SelectedTablesForAudit {
    $selected = [System.Collections.Generic.List[string]]::new()

    if ($cmbBeforeTable.SelectedItem -and $cmbBeforeTable.SelectedItem -ne "(Todas)") {
        $value = [string]$cmbBeforeTable.SelectedItem
        if (-not $selected.Contains($value)) {
            [void]$selected.Add($value)
        }
    }

    if ($cmbAfterTable.SelectedItem -and $cmbAfterTable.SelectedItem -ne "(Todas)") {
        $value = [string]$cmbAfterTable.SelectedItem
        if (-not $selected.Contains($value)) {
            [void]$selected.Add($value)
        }
    }

    return [string[]]@($selected)
}

function Bind-TableGrid {
    param(
        [object[]]$TableReports
    )

    $gridRows = foreach ($table in $TableReports) {
        [pscustomobject]@{
            Tabla = $table.table
            Altas = $table.inserted_count
            Bajas = $table.deleted_count
            Modificados = $table.modified_count
            Clave = $(if ($table.has_reliable_key) { ($table.primary_key_columns -join ", ") } else { "Sin PK fiable" })
            Nota = $(if ($table.PSObject.Properties.Name -contains "note") { $table.note } else { "" })
        }
    }

    $gridTables.DataSource = $null
    $gridTables.DataSource = $gridRows
}

function Reset-DetailViews {
    $listInserted.Items.Clear()
    $listDeleted.Items.Clear()
    $listModified.Items.Clear()
    $txtInserted.Clear()
    $txtDeleted.Clear()
    $txtModified.Clear()
    $currentInserted = @()
    $currentDeleted = @()
    $currentModified = @()
}

function Bind-TableDetails {
    param(
        [object]$TableReport
    )

    Reset-DetailViews

    $script:currentInserted = @($TableReport.inserted)
    $script:currentDeleted = @($TableReport.deleted)
    $script:currentModified = @($TableReport.modified)

    foreach ($item in $script:currentInserted) {
        [void]$listInserted.Items.Add($item.key)
    }
    foreach ($item in $script:currentDeleted) {
        [void]$listDeleted.Items.Add($item.key)
    }
    foreach ($item in $script:currentModified) {
        [void]$listModified.Items.Add($item.key)
    }

    if ($listInserted.Items.Count -gt 0) { $listInserted.SelectedIndex = 0 }
    if ($listDeleted.Items.Count -gt 0) { $listDeleted.SelectedIndex = 0 }
    if ($listModified.Items.Count -gt 0) { $listModified.SelectedIndex = 0 }
}

$btnBefore.Add_Click({
    $selectedFile = Choose-AccessFile
    if ($selectedFile) {
        $txtBefore.Text = $selectedFile
    }
})

$btnAfter.Add_Click({
    $selectedFile = Choose-AccessFile
    if ($selectedFile) {
        $txtAfter.Text = $selectedFile
    }
})

$btnLoadTables.Add_Click({
    try {
        Load-TableLists
    }
    catch {
        Show-ErrorMessage $_.Exception.Message
        $statusLabel.Text = "Error cargando tablas."
    }
})

$btnAudit.Add_Click({
    try {
        $statusLabel.Text = "Ejecutando auditoria..."
        $form.UseWaitCursor = $true
        $form.Refresh()

        $selectedTables = Get-SelectedTablesForAudit
        $ignoreColumns = @(
            $txtIgnore.Text.Split(",", [System.StringSplitOptions]::RemoveEmptyEntries) |
                ForEach-Object { $_.Trim() } |
                Where-Object { $_ }
        )

        $invokeParams = @{
            BeforePath = $txtBefore.Text
            AfterPath = $txtAfter.Text
            IncludeUnchangedTables = $chkUnchanged.Checked
        }

        if (@($ignoreColumns).Count -gt 0) {
            $invokeParams.IgnoreColumns = $ignoreColumns
        }

        if (@($selectedTables).Count -gt 0) {
            $invokeParams.TableNames = $selectedTables
        }

        $script:currentResult = Invoke-AccessAuditComparison @invokeParams

        $script:currentTableReports = @($script:currentResult.tables)
        Bind-TableGrid -TableReports $script:currentTableReports
        Reset-DetailViews
        $btnExport.Enabled = $true

        if (@($script:currentTableReports).Count -eq 0) {
            $statusLabel.Text = "No se detectaron tablas con cambios."
        }
        else {
            $statusLabel.Text = "Auditoria completada. Selecciona una tabla para ver detalles."
        }
    }
    catch {
        Show-ErrorMessage $_.Exception.Message
        $statusLabel.Text = "La auditoria ha fallado."
    }
    finally {
        $form.UseWaitCursor = $false
    }
})

$gridTables.Add_SelectionChanged({
    if (-not $gridTables.SelectedRows -or $gridTables.SelectedRows.Count -eq 0) {
        return
    }

    $selectedTableName = [string]$gridTables.SelectedRows[0].Cells["Tabla"].Value
    $tableReport = $script:currentTableReports | Where-Object { $_.table -eq $selectedTableName } | Select-Object -First 1
    if ($tableReport) {
        Bind-TableDetails -TableReport $tableReport
    }
})

$listInserted.Add_SelectedIndexChanged({
    if ($listInserted.SelectedIndex -ge 0) {
        $txtInserted.Text = Format-RowDetails -Entry $script:currentInserted[$listInserted.SelectedIndex] -Mode "inserted"
    }
})

$listDeleted.Add_SelectedIndexChanged({
    if ($listDeleted.SelectedIndex -ge 0) {
        $txtDeleted.Text = Format-RowDetails -Entry $script:currentDeleted[$listDeleted.SelectedIndex] -Mode "deleted"
    }
})

$listModified.Add_SelectedIndexChanged({
    if ($listModified.SelectedIndex -ge 0) {
        $txtModified.Text = Format-RowDetails -Entry $script:currentModified[$listModified.SelectedIndex] -Mode "modified"
    }
})

$btnExport.Add_Click({
    try {
        if (-not $script:currentResult) {
            throw "Todavia no hay una auditoria ejecutada."
        }

        $dialog = [System.Windows.Forms.SaveFileDialog]::new()
        $dialog.Filter = "JSON (*.json)|*.json"
        $dialog.FileName = "audit-report.json"
        if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            return
        }

        $script:currentResult | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $dialog.FileName -Encoding UTF8
        Show-InfoMessage "Informe exportado a:`n$($dialog.FileName)"
        $statusLabel.Text = "Informe exportado correctamente."
    }
    catch {
        Show-ErrorMessage $_.Exception.Message
        $statusLabel.Text = "No se pudo exportar el informe."
    }
})

[void]$form.ShowDialog()
