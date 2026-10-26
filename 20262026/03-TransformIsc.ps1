<#
.SYNOPSIS
  Stage 3 - transform-isc candidates from inventory. Analyze output is optional.

.DESCRIPTION
  With analyze (default when 02-analyze has object-analyses.json or manifest.json):
    honors MIGRATE / REDESIGN / QUARANTINE / EXCLUDE and finding-resolutions.json.

  Without analyze (-SkipAnalyze, or analyze folder empty):
    transform-isc is permissive and maps all Applications in the inventory.

.EXAMPLE
  .\03-TransformIsc.ps1 -WorkRoot C:\work\iiq-isc-run

.EXAMPLE
  .\03-TransformIsc.ps1 -WorkRoot C:\work\iiq-isc-run -SkipAnalyze

.EXAMPLE
  .\03-TransformIsc.ps1 -WorkRoot C:\work\iiq-isc-run -JdbcNameSuffix "-JDBC" -DelimitedNameSuffix "-DELIMITED"
#>
param(
  [string]$WorkRoot,
  [string]$RunId,
  [string]$AttemptId,
  [string]$Source,
  [string]$Mapping,
  [string]$ExecutionGate,
  [string]$ApplicationType,
  [string]$ApplicationNames,
  [string]$ExcludeApplicationNames,
  [string]$SqlLoaderNameSuffix,
  # Optional. Empty by default — only passed to ezy-transfer when a value is set.
  [string]$JdbcNameSuffix,
  [string]$DelimitedNameSuffix,
  [switch]$SkipAnalyze,
  [switch]$RequireAnalyze
)

. (Join-Path $PSScriptRoot "Common.ps1")

$context = Ensure-EzyRunContext `
  -WorkRoot $WorkRoot `
  -RunId $RunId `
  -AttemptId $AttemptId `
  -Source $Source `
  -Mapping $Mapping `
  -ExecutionGate $ExecutionGate `
  -ApplicationType $ApplicationType `
  -ApplicationNames $ApplicationNames `
  -ExcludeApplicationNames $ExcludeApplicationNames

Write-EzyBanner -Stage "transform-isc" -Context $context

$inventory = Get-EzyInventoryRoot -Context $context
$analyze = [string]$context.dirs.analyze
$out = [string]$context.dirs.transform
New-Item -ItemType Directory -Force -Path $out | Out-Null

$analyzePresent = (Test-Path -LiteralPath (Join-Path $analyze "object-analyses.json")) `
  -or (Test-Path -LiteralPath (Join-Path $analyze "manifest.json"))
$useAnalyze = -not $SkipAnalyze -and $analyzePresent

if ($RequireAnalyze -and -not $useAnalyze) {
  throw "Analyze output missing under $analyze. Run .\02-Analyze.ps1 or omit -RequireAnalyze."
}

$cli = New-Object System.Collections.Generic.List[string]
$cli.Add("transform-isc")
$cli.Add("--source"); $cli.Add($inventory)
$cli.Add("--out"); $cli.Add($out)
if ($useAnalyze) {
  $cli.Add("--analysis"); $cli.Add($analyze)
  Write-Host ("analysis={0}" -f $analyze)
} else {
  Write-Host "analysis=(none) - transform-isc will map all Applications (no disposition gate)"
  if ($SkipAnalyze) {
    Write-Host "SkipAnalyze was set."
  } else {
    Write-Host "No 02-analyze\object-analyses.json or manifest.json found. Run .\02-Analyze.ps1 first if you want holds."
  }
  $provisioningFile = Join-Path $analyze "provisioning-capabilities.json"
  if (Test-Path -LiteralPath $provisioningFile) {
    $cli.Add("--provisioning"); $cli.Add($analyze)
    Write-Host ("provisioning={0}" -f $analyze)
  }
}
$cli.Add("--mapping"); $cli.Add([string]$context.mapping)
$cli.Add("--execution-gate"); $cli.Add([string]$context.executionGate)
if (-not $PSBoundParameters.ContainsKey("SqlLoaderNameSuffix")) {
  if ($null -ne $env:EZYMIGRATE_SQLLOADER_NAME_SUFFIX) {
    $SqlLoaderNameSuffix = $env:EZYMIGRATE_SQLLOADER_NAME_SUFFIX
  } else {
    $SqlLoaderNameSuffix = "-SQLLOADER"
  }
}
$cli.Add("--sqlloader-name-suffix"); $cli.Add([string]$SqlLoaderNameSuffix)
Write-Host ("sqlLoaderNameSuffix={0}" -f $SqlLoaderNameSuffix)

# JDBC / Delimited suffixes: no default. Only add CLI flags when a value is provided
# (param or env). Omit the flag to leave ezy-transfer naming unchanged for that type.
if (-not $PSBoundParameters.ContainsKey("JdbcNameSuffix") `
    -and -not [string]::IsNullOrWhiteSpace($env:EZYMIGRATE_JDBC_NAME_SUFFIX)) {
  $JdbcNameSuffix = $env:EZYMIGRATE_JDBC_NAME_SUFFIX
}
if (-not [string]::IsNullOrWhiteSpace($JdbcNameSuffix)) {
  $cli.Add("--jdbc-name-suffix"); $cli.Add($JdbcNameSuffix.Trim())
  Write-Host ("jdbcNameSuffix={0}" -f $JdbcNameSuffix.Trim())
} else {
  Write-Host "jdbcNameSuffix=(none)"
}

if (-not $PSBoundParameters.ContainsKey("DelimitedNameSuffix") `
    -and -not [string]::IsNullOrWhiteSpace($env:EZYMIGRATE_DELIMITED_NAME_SUFFIX)) {
  $DelimitedNameSuffix = $env:EZYMIGRATE_DELIMITED_NAME_SUFFIX
}
if (-not [string]::IsNullOrWhiteSpace($DelimitedNameSuffix)) {
  $cli.Add("--delimited-name-suffix"); $cli.Add($DelimitedNameSuffix.Trim())
  Write-Host ("delimitedNameSuffix={0}" -f $DelimitedNameSuffix.Trim())
} else {
  Write-Host "delimitedNameSuffix=(none)"
}

$cli.Add("--run-id"); $cli.Add([string]$context.runId)
$cli.Add("--attempt-id"); $cli.Add([string]$context.attemptId)
Add-EzyFilterArgs -ArgsList $cli -Context $context

Invoke-EzyTransfer -CliArgs $cli.ToArray()
Set-EzyLastStage -Context $context -Stage "transform-isc"
Write-Host "Next: .\04-Validate.ps1"
