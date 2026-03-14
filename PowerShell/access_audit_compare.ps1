param(
    [Parameter(Mandatory = $true)]
    [string]$BeforePath,

    [Parameter(Mandatory = $true)]
    [string]$AfterPath,

    [string]$OutputPath = ".\\audit-report.json",

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

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. "$PSScriptRoot\\AccessAuditCore.ps1"

$invokeParams = @{
    BeforePath = $BeforePath
    AfterPath = $AfterPath
    IncludeUnchangedTables = $IncludeUnchangedTables
}

if ($null -ne $IgnoreColumns -and @($IgnoreColumns).Count -gt 0) {
    $invokeParams.IgnoreColumns = $IgnoreColumns
}

if ($null -ne $TableNames -and @($TableNames).Count -gt 0) {
    $invokeParams.TableNames = $TableNames
}

$result = Invoke-AccessAuditComparison @invokeParams

$outputDirectory = Split-Path -Path $OutputPath -Parent
if ($outputDirectory -and -not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}

$result | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
Write-ConsoleSummary -TableReports $result.tables
Write-Host ""
Write-Host "Informe JSON generado en: $((Resolve-Path -LiteralPath $OutputPath).Path)"
