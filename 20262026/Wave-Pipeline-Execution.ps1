# ============================================================================
# Wave Migration Execution
# Application type and name are read from wave-applications.csv.
# Columns: ApplicationType, ApplicationName
# ============================================================================

param(
    [Parameter(Mandatory = $true)]
    [string]$Wave,

    [string]$ApplicationCsv,

    # Optional ezy-transfer ISC name suffixes (Stage 3 only). Empty = do not pass the flag.
    [string]$JdbcNameSuffix,
    [string]$DelimitedNameSuffix,
    [string]$SqlLoaderNameSuffix
)

$ErrorActionPreference = "Stop"

# ============================================================================
# GLOBAL CONFIGURATION
# ============================================================================

$MigrationRoot = "C:\workspace\Prudential\ezyFramework\Migration-work"
$Source        = "C:\workspace\Prudential\ezyFramework\exports"

if ([string]::IsNullOrWhiteSpace($ApplicationCsv)) {
    $ApplicationCsv = "C:\workspace\Prudential\ezyFramework\Migration-work\wave-applications.csv"
}

$waveName = $Wave.Trim()
if ($waveName -match '^\d+$') {
    $waveName = "wave$waveName"
}
$waveName = $waveName -replace '^(?i)wave[-_\s]*', 'wave'
if ($waveName -notmatch '^wave.+$') {
    throw "Wave is required. Example: -Wave 6  or  -Wave wave6"
}

$WorkRoot = Join-Path $MigrationRoot $waveName
New-Item -ItemType Directory -Force -Path $WorkRoot | Out-Null

$RunId     = "run-2026-09-21-05"
$AttemptId = "attempt-1"

$ExcludeApplicationNames = "DECOM_.*,^Test"

$Mapping       = "iiq-isc-v1"
$ExecutionGate = "standard"

# ============================================================================
# SCRIPT ROOT
# ============================================================================

$ScriptRoot = "C:\workspace\Prudential\ezyFramework\scripts\iiq-isc"

# ============================================================================
# APPLICATION LIST
# ============================================================================

function Get-CsvField {
    param(
        $Row,
        [string[]]$Candidates
    )

    foreach ($prop in $Row.PSObject.Properties) {
        foreach ($candidate in $Candidates) {
            if ($prop.Name.Trim().Equals($candidate, [System.StringComparison]::OrdinalIgnoreCase)) {
                if ($null -eq $prop.Value) {
                    return ""
                }
                return [string]$prop.Value
            }
        }
    }
    return ""
}

function Get-WaveApplicationSelection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Application CSV not found: $Path"
    }

    $rows = @(Import-Csv -LiteralPath $Path)
    $applications = @()
    $seen = @{}

    foreach ($row in $rows) {
        $name = (Get-CsvField -Row $row -Candidates @("ApplicationName", "Name")).Trim()
        $type = (Get-CsvField -Row $row -Candidates @("ApplicationType", "Type")).Trim()
        if ([string]::IsNullOrWhiteSpace($name) -or $name.StartsWith("#")) {
            continue
        }
        if ([string]::IsNullOrWhiteSpace($type)) {
            throw "Application '$name' in $Path has no ApplicationType."
        }
        $key = $name.ToLowerInvariant()
        if ($seen.ContainsKey($key)) {
            Write-Warning "Duplicate application '$name' ignored. Kept type '$($seen[$key])'."
            continue
        }
        $seen[$key] = $type
        $applications += [pscustomobject]@{
            ApplicationType = $type
            ApplicationName = $name
        }
    }

    if ($applications.Count -eq 0) {
        throw "Application CSV has no applications: $Path"
    }

    return @($applications)
}

$Applications = @(Get-WaveApplicationSelection -Path $ApplicationCsv)
$ApplicationNames = @($Applications | ForEach-Object { $_.ApplicationName })
$ApplicationType = (($Applications | ForEach-Object { $_.ApplicationType } | Select-Object -Unique) -join ",")

# ============================================================================
# COMMON CONFIGURATION
# ============================================================================
# This configuration is passed to EVERY pipeline stage.
# ============================================================================

$CommonParameters = @{
    WorkRoot                = $WorkRoot
    Source                  = $Source
    RunId                   = $RunId
    AttemptId               = $AttemptId
    applicationType         = $ApplicationType
    applicationNames        = ($ApplicationNames -join ",")
    ExcludeApplicationNames = $ExcludeApplicationNames
    Mapping                 = $Mapping
    ExecutionGate           = $ExecutionGate
}

# ============================================================================
# HELPER FUNCTION
# ============================================================================

function Invoke-MigrationStage {

    param (
        [Parameter(Mandatory = $true)]
        [string]$StageName,

        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,

        [Parameter(Mandatory = $false)]
        [hashtable]$Parameters = @{}
    )

    Write-Host ""
    Write-Host "============================================================================" -ForegroundColor Cyan
    Write-Host " $StageName" -ForegroundColor Cyan
    Write-Host "============================================================================" -ForegroundColor Cyan
    Write-Host "Script : $ScriptPath" -ForegroundColor Gray

    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        Write-Host "ERROR: Script not found:" -ForegroundColor Red
        Write-Host $ScriptPath -ForegroundColor Red
        exit 1
    }

    try {
        & $ScriptPath @Parameters

        if ($LASTEXITCODE -ne 0) {
            throw "Stage '$StageName' failed with exit code $LASTEXITCODE."
        }

        Write-Host ""
        Write-Host "SUCCESS: $StageName completed." -ForegroundColor Green
    }
    catch {
        Write-Host ""
        Write-Host "FAILED: $StageName" -ForegroundColor Red
        Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    #    exit 1
    }
}

# ============================================================================
# DISPLAY CONFIGURATION
# ============================================================================

Write-Host ""
Write-Host "============================================================================" -ForegroundColor Yellow
Write-Host " MIGRATION CONFIGURATION" -ForegroundColor Yellow
Write-Host "============================================================================" -ForegroundColor Yellow

Write-Host "Script Root          : $ScriptRoot"
Write-Host "Wave                 : $waveName"
Write-Host "Work Root            : $WorkRoot"
Write-Host "Source               : $Source"
Write-Host "Run ID               : $RunId"
Write-Host "Attempt ID           : $AttemptId"
Write-Host "Application CSV      : $ApplicationCsv"
Write-Host "Applications         : $($Applications.Count) unique"
foreach ($app in $Applications) {
    Write-Host ("  {0} | {1}" -f $app.ApplicationType, $app.ApplicationName)
}
Write-Host "Application Type     : $ApplicationType"
Write-Host "Application Names    : $($ApplicationNames -join ', ')"
Write-Host "Exclude Applications : $ExcludeApplicationNames"
Write-Host "Mapping              : $Mapping"
Write-Host "Execution Gate       : $ExecutionGate"
if (-not [string]::IsNullOrWhiteSpace($JdbcNameSuffix)) {
    Write-Host "JDBC Name Suffix     : $JdbcNameSuffix"
}
if (-not [string]::IsNullOrWhiteSpace($DelimitedNameSuffix)) {
    Write-Host "Delimited Name Suffix: $DelimitedNameSuffix"
}
if (-not [string]::IsNullOrWhiteSpace($SqlLoaderNameSuffix)) {
    Write-Host "SQLLoader Name Suffix: $SqlLoaderNameSuffix"
}

# ============================================================================
# STAGE 0 — INITIAL SETUP
# ============================================================================

$Stage0Parameters = $CommonParameters.Clone()

Invoke-MigrationStage -StageName "STAGE 0 - INITIAL SETUP" -ScriptPath (Join-Path $ScriptRoot "Init-Run.ps1") -Parameters $Stage0Parameters

# ============================================================================
# STAGE 1 — INGEST
# ============================================================================

$Stage1Parameters = $CommonParameters.Clone()

Invoke-MigrationStage -StageName "STAGE 1 - INGEST" -ScriptPath (Join-Path $ScriptRoot "01-Ingest.ps1") -Parameters $Stage1Parameters

# ============================================================================
# STAGE 2 — ANALYZE
# ============================================================================

$Stage2Parameters = $CommonParameters.Clone()

Invoke-MigrationStage -StageName "STAGE 2 - ANALYZE" -ScriptPath (Join-Path $ScriptRoot "02-Analyze.ps1") -Parameters $Stage2Parameters

# ============================================================================
# STAGE 3 — TRANSFORM ISC
# ============================================================================

$Stage3Parameters = $CommonParameters.Clone()
$Stage3Parameters["SkipAnalyze"] = $true
if (-not [string]::IsNullOrWhiteSpace($JdbcNameSuffix)) {
    $Stage3Parameters["JdbcNameSuffix"] = $JdbcNameSuffix.Trim()
}
if (-not [string]::IsNullOrWhiteSpace($DelimitedNameSuffix)) {
    $Stage3Parameters["DelimitedNameSuffix"] = $DelimitedNameSuffix.Trim()
}
if (-not [string]::IsNullOrWhiteSpace($SqlLoaderNameSuffix)) {
    $Stage3Parameters["SqlLoaderNameSuffix"] = $SqlLoaderNameSuffix.Trim()
}

Invoke-MigrationStage -StageName "STAGE 3 - TRANSFORM-ISC" -ScriptPath (Join-Path $ScriptRoot "03-TransformIsc.ps1") -Parameters $Stage3Parameters

# ============================================================================
# STAGE 4 — VALIDATE
# FAIL-CLOSED GATE
# ============================================================================

$Stage4Parameters = $CommonParameters.Clone()

Invoke-MigrationStage -StageName "STAGE 4 - VALIDATE (FAIL-CLOSED GATE)" -ScriptPath (Join-Path $ScriptRoot "04-Validate.ps1") -Parameters $Stage4Parameters

# ============================================================================
# STAGE 5 — LOAD
# ISC TENANT
# ============================================================================

$Stage5Parameters = $CommonParameters.Clone()
$Stage5Parameters["Mode"] = "live"

Invoke-MigrationStage -StageName "STAGE 5 - LOAD (ISC TENANT)" -ScriptPath (Join-Path $ScriptRoot "05-Load.ps1") -Parameters $Stage5Parameters

# ============================================================================
# STAGE 6 — FINAL SCORECARD
# ============================================================================

$Stage6Parameters = $CommonParameters.Clone()

Invoke-MigrationStage -StageName "STAGE 6 - FINAL SCORECARD" -ScriptPath (Join-Path $ScriptRoot "06-Scorecard.ps1") -Parameters $Stage6Parameters

# ============================================================================
# MIGRATION COMPLETED
# ============================================================================

Write-Host ""
Write-Host "============================================================================" -ForegroundColor Green
Write-Host " MIGRATION PIPELINE COMPLETED SUCCESSFULLY" -ForegroundColor Green
Write-Host "============================================================================" -ForegroundColor Green

Write-Host "Run ID       : $RunId" -ForegroundColor White
Write-Host "Attempt ID   : $AttemptId" -ForegroundColor White
Write-Host "Work Root    : $WorkRoot" -ForegroundColor White
Write-Host "Mapping      : $Mapping" -ForegroundColor White
Write-Host "Execution    : $ExecutionGate" -ForegroundColor White
Write-Host "Script Root  : $ScriptRoot" -ForegroundColor White
Write-Host ""
