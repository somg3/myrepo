#Requires -Version 5.1
<#
.SYNOPSIS
    Uploads changed source CSV files to SailPoint ISC with the File Upload Utility.

.DESCRIPTION
    Watches a source-input folder for CSV files. The file name must match exactly
    one SailPoint source name. Account feeds are named {SourceName}.csv. Group feeds
    are named {SourceName}-group.csv. Unchanged files are skipped by SHA-256
    checksum, tracked separately for account and group. Each changed file is copied
    to a temporary name that starts with the source ID, then submitted with
    sailpoint-file-upload-utility. Account files use account aggregation. Group
    files use entitlement aggregation with schema name group. When both are
    selected, group files are aggregated first and account files follow.

    Credentials are never read from this script and are never placed on the Java
    command line. Supply them with SAIL_CLIENT_ID and SAIL_CLIENT_SECRET, or save
    them with -SaveCredential (DPAPI, current Windows user).

    Non-secret settings can live in upload-config.json under the utility folder.
    Parameter values override environment variables. Environment variables override
    the config file.

.PARAMETER UtilityDir
    Root folder for the jar, config, logs, state, and lock file.
    Default: the folder that contains this script, so the same copy works on C:, D:, F:, or any other drive.
    Override with this parameter or with SAILPOINT_UTILITY_DIR.
    Example: F:\SailPointUtility\FileUploadUtility.

.PARAMETER SourceDir
    Folder that contains the source CSV files. A full path on any drive is accepted,
    for example D:\Feeds\Source Input. A name without a drive is created under UtilityDir.
    Config key sourceDirectory. Environment variable SAILPOINT_SOURCE_DIR.

.PARAMETER TemporaryDir
    Folder for the short-lived upload copies. A full path on any drive is accepted,
    for example E:\SailPoint\Temporary. A name without a drive is created under UtilityDir.
    Config key temporaryDirectory. Environment variable SAILPOINT_TEMPORARY_DIR.

.PARAMETER TenantUrl
    SailPoint API base URL, for example https://tenant.api.identitynow.com.
    Falls back to SAIL_BASE_URL, then SAILPOINT_TENANT_URL, then config tenantUrl.

.PARAMETER ConfigPath
    Path to upload-config.json. Default: <UtilityDir>\upload-config.json.

.PARAMETER JarName
    File Upload Utility jar name inside UtilityDir.

.PARAMETER SourcesPageSize
    Page size for GET /v3/sources. Range 1-250.

.PARAMETER FileStabilityWaitSeconds
    Seconds between two size/timestamp samples before a CSV is accepted.

.PARAMETER RunsToKeep
    Number of Run_* log folders to retain.

.PARAMETER PauseAtEnd
    Wait for ENTER before exiting. Ignored when there is no interactive console.

.PARAMETER EnableSourceDiscoveryLogging
    Write every discovered source name and ID to the summary log.

.PARAMETER CaseSensitiveSourceMatch
    Match CSV names to source names with ordinal case sensitivity.
    The default match is ordinal and case-insensitive, same as the previous script.

.PARAMETER JavaTimeoutMinutes
    Maximum time allowed for one Java upload. Range 1-240.

.PARAMETER ApiRetryCount
    Attempts for retryable SailPoint API failures (408, 429, 5xx, network). Range 1-5.

.PARAMETER ObjectType
    Which feeds to aggregate: account, group, or both.
    Account files are {SourceName}.csv. Group files are {SourceName}-group.csv.
    Config key objectType. Environment variable SAILPOINT_OBJECT_TYPE.
    Default: account. A parameter overrides the environment variable, which overrides the config file.
    When both is selected, group files are aggregated first and account files follow.

.PARAMETER SaveCredential
    Prompt for the client ID and secret, store them with DPAPI for the current user,
    and exit. Run this once as the scheduled-task account.

.PARAMETER WhatIf
    Resolve sources and checksums, and do not upload or update success state.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Sailpoint_ISC_File_Upload_Utility.ps1 -SaveCredential

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Sailpoint_ISC_File_Upload_Utility.ps1 -TenantUrl "https://example.api.identitynow.com"

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Sailpoint_ISC_File_Upload_Utility.ps1 -WhatIf

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Sailpoint_ISC_File_Upload_Utility.ps1 -ObjectType both

.NOTES
    Exit codes:
      0  All eligible files were processed, or no CSV files were present.
      1  Startup failure, or one or more file failures.
      3  Another instance holds the execution lock.
      4  Uploads succeeded, but cleanup reported a warning.

    Each run writes three logs under the run folder:
      summary_log_*.txt       every INFO, SUCCESS, WARNING, and ERROR line
      error_log_*.txt         WARNING and ERROR lines only
      *_aggregation_log_*.txt one file for each CSV that was considered.
        The log name follows the input file name, so group and account feeds stay distinct:
          HR-APP-SG.csv       -> HR-APP-SG_aggregation_log_<stamp>.txt
          HR-APP-SG-Group.csv -> HR-APP-SG-Group_aggregation_log_<stamp>.txt

    Startup failures are written to those logs as soon as the utility folder
    is known, including config, credential, Java, and SailPoint API errors.

    Config file keys (no secrets):
      tenantUrl, jarName, sourcesPageSize, fileStabilityWaitSeconds, runsToKeep,
      enableSourceDiscoveryLogging, pauseAtEnd, caseSensitiveSourceMatch,
      javaTimeoutMinutes, apiRetryCount, objectType, sourceDirectory, temporaryDirectory,
      inputFolderName, processingFolderName, logFolderName, stateFolderName

    sourceDirectory and temporaryDirectory accept a full path on any drive.
    inputFolderName and processingFolderName remain folder names under UtilityDir
    when the full path is not set. The default temporary folder name remains
    "Temperary Files" so existing folders keep working.

    Rotate any client secret that was previously stored in an older copy of this
    script. That credential must be treated as exposed.
#>
[CmdletBinding(SupportsShouldProcess = $false)]
param(
    [string]$UtilityDir,
    [string]$SourceDir,
    [string]$TemporaryDir,
    [string]$TenantUrl,
    [string]$ConfigPath,
    [string]$JarName,
    [int]$SourcesPageSize,
    [int]$FileStabilityWaitSeconds,
    [int]$RunsToKeep,
    [switch]$PauseAtEnd,
    [switch]$EnableSourceDiscoveryLogging,
    [switch]$CaseSensitiveSourceMatch,
    [int]$JavaTimeoutMinutes,
    [int]$ApiRetryCount,
    [string]$ObjectType,
    [switch]$SaveCredential,
    [switch]$WhatIf
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ScriptVersion = '5.3.3'
$script:RequestedExitCode = 1
$script:StopRequested = $false
$script:LockStream = $null
$script:LockHeld = $false
$script:LockPath = $null
$script:RunSummaryLog = $null
$script:RunErrorLog = $null
$script:RunLogDir = $null
$script:RunStarted = $null
$script:RunId = $null
$script:RunStamp = $null
$script:UtilityPath = $null
$script:SummaryFinalized = $false
$script:OpeningLog = $false
$script:PendingSummaryLines = New-Object 'System.Collections.Generic.List[string]'
$script:PendingErrorLines = New-Object 'System.Collections.Generic.List[string]'
$script:CredentialSource = 'not resolved'
$script:JavaHomeNote = $null
$script:RedactedValues = New-Object 'System.Collections.Generic.List[string]'
$script:ActiveJavaProcess = $null
$script:ClientId = $null
$script:ClientSecret = $null

$script:SuccessCount = 0
$script:FailureCount = 0
$script:SkippedCount = 0
$script:UnchangedCount = 0
$script:WhatIfCount = 0
$script:CleanupWarningCount = 0


function Initialize-JavaEnvironment {
    $existingHome = [Environment]::GetEnvironmentVariable('JAVA_HOME')
    if (-not [string]::IsNullOrWhiteSpace($existingHome)) {
        $existingJava = Join-Path $existingHome.Trim() 'bin\java.exe'
        if (Test-Path -LiteralPath $existingJava) {
            $env:JAVA_HOME = $existingHome.Trim()
            $env:Path = "$(Split-Path -Parent $existingJava);$env:Path"
            return
        }

        $script:JavaHomeNote = "JAVA_HOME is set but java.exe was not found at $existingJava. The script will look for Java on PATH."
    }

    $knownJava = 'C:\Program Files\Java\jdk-21.0.11\bin\java.exe'
    if (Test-Path -LiteralPath $knownJava) {
        $env:JAVA_HOME = Split-Path -Parent (Split-Path -Parent $knownJava)
        $env:Path = "$(Split-Path -Parent $knownJava);$env:Path"
    }
}

Initialize-JavaEnvironment

function Enable-Tls12 {
    try {
        $current = [Net.ServicePointManager]::SecurityProtocol
        $tls12 = [Net.SecurityProtocolType]::Tls12
        if (($current -band $tls12) -ne $tls12) {
            [Net.ServicePointManager]::SecurityProtocol = $current -bor $tls12
        }
    }
    catch {
        Write-Warning "TLS 1.2 could not be enabled explicitly. $($_.Exception.Message)"
    }
}

function ConvertTo-SingleLine {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrEmpty($Value)) {
        return ''
    }

    return (($Value -replace '[\r\n]+', ' ') -replace '\s{2,}', ' ').Trim()
}

function Hide-SecretText {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrEmpty($Value) -or $script:RedactedValues.Count -eq 0) {
        return $Value
    }

    $redacted = $Value
    foreach ($secret in $script:RedactedValues) {
        if ($secret -and $secret.Length -ge 8 -and $redacted.Contains($secret)) {
            $redacted = $redacted.Replace($secret, '***REDACTED***')
        }
    }

    return $redacted
}

function Register-RedactedValue {
    param([AllowNull()][string]$Value)

    if ($Value -and $Value.Length -ge 8 -and -not $script:RedactedValues.Contains($Value)) {
        $script:RedactedValues.Add($Value)
    }
}

function Write-Utf8Lines {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][AllowEmptyString()][AllowEmptyCollection()][string[]]$Lines,
        [switch]$Append
    )

    $encoding = New-Object System.Text.UTF8Encoding $false
    if ($Append) {
        [System.IO.File]::AppendAllLines($Path, $Lines, $encoding)
    }
    else {
        [System.IO.File]::WriteAllLines($Path, $Lines, $encoding)
    }
}

function Write-RunLog {
    param(
        [AllowNull()][string]$Message,
        [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR')]
        [string]$Level = 'INFO'
    )

    try {
        $safeMessage = ConvertTo-SingleLine -Value (Hide-SecretText -Value $Message)
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
        $entry = "[$timestamp] [$Level] $safeMessage"

        if ($script:RunSummaryLog) {
            Write-Utf8Lines -Path $script:RunSummaryLog -Lines @($entry) -Append
        }
        else {
            $script:PendingSummaryLines.Add($entry)
        }

        if ($Level -eq 'WARNING' -or $Level -eq 'ERROR') {
            if ($script:RunErrorLog) {
                Write-Utf8Lines -Path $script:RunErrorLog -Lines @($entry) -Append
            }
            else {
                $script:PendingErrorLines.Add($entry)
            }
        }

        $color = switch ($Level) {
            'SUCCESS' { 'Green' }
            'WARNING' { 'Yellow' }
            'ERROR' { 'Red' }
            default { 'Gray' }
        }

        Write-Host $entry -ForegroundColor $color
    }
    catch {
        Write-Host "Logging failure: $(ConvertTo-SingleLine $_.Exception.Message)" -ForegroundColor Yellow
    }
}

function Format-ErrorDetail {
    param($ErrorRecord)

    if ($null -eq $ErrorRecord) {
        return 'Unknown error.'
    }

    $parts = New-Object System.Collections.Generic.List[string]
    $exception = $ErrorRecord.Exception
    $depth = 0
    while ($null -ne $exception -and $depth -lt 6) {
        $typeName = $exception.GetType().Name
        $text = ConvertTo-SingleLine ([string]$exception.Message)
        if ([string]::IsNullOrWhiteSpace($text)) {
            $text = $typeName
        }
        $parts.Add("${typeName}: $text")
        $exception = $exception.InnerException
        $depth++
    }

    $location = ''
    if ($ErrorRecord.InvocationInfo -and $ErrorRecord.InvocationInfo.ScriptLineNumber -gt 0) {
        $location = " ScriptLine=$($ErrorRecord.InvocationInfo.ScriptLineNumber)"
    }

    $category = ''
    if ($ErrorRecord.CategoryInfo -and $ErrorRecord.CategoryInfo.Category) {
        $category = " Category=$($ErrorRecord.CategoryInfo.Category)"
    }

    return (Hide-SecretText (($parts -join ' | ') + $category + $location)).Trim()
}

function Open-RunLogs {
    param([Parameter(Mandatory = $true)][string]$LogRoot)

    if ($script:RunSummaryLog -or $script:OpeningLog) {
        return
    }

    $script:OpeningLog = $true
    try {
        if (-not (Test-Path -LiteralPath $LogRoot -PathType Container)) {
            New-Item -Path $LogRoot -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }

        $script:RunStarted = Get-Date
        $script:RunId = [guid]::NewGuid().ToString('n')
        $script:RunStamp = Get-RunStamp
        $script:RunLogDir = Join-Path $LogRoot "Run_$($script:RunStamp)"
        New-Item -Path $script:RunLogDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        $script:RunSummaryLog = Join-Path $script:RunLogDir "summary_log_$($script:RunStamp).txt"
        $script:RunErrorLog = Join-Path $script:RunLogDir "error_log_$($script:RunStamp).txt"

        Write-Utf8Lines -Path $script:RunSummaryLog -Lines @(
            '======================================================================'
            'SAILPOINT ISC MULTI-SOURCE UPLOAD SUMMARY'
            '======================================================================'
            "Script version    : $script:ScriptVersion"
            "Run ID            : $($script:RunId)"
            "Run started       : $($script:RunStarted.ToString('yyyy-MM-dd HH:mm:ss.fff'))"
            "Computer          : $env:COMPUTERNAME"
            "Executed by       : $env:USERNAME"
            "PowerShell PID    : $PID"
            '======================================================================'
            ''
        )

        Write-Utf8Lines -Path $script:RunErrorLog -Lines @(
            '======================================================================'
            'SAILPOINT ISC UPLOAD ERROR LOG'
            '======================================================================'
            "Script version    : $script:ScriptVersion"
            "Run ID            : $($script:RunId)"
            'This file contains WARNING and ERROR entries only.'
            'The summary log contains the full trace.'
            '======================================================================'
            ''
        )

        if ($script:PendingSummaryLines.Count -gt 0) {
            Write-Utf8Lines -Path $script:RunSummaryLog -Lines @($script:PendingSummaryLines.ToArray()) -Append
            $script:PendingSummaryLines.Clear()
        }

        if ($script:PendingErrorLines.Count -gt 0) {
            Write-Utf8Lines -Path $script:RunErrorLog -Lines @($script:PendingErrorLines.ToArray()) -Append
            $script:PendingErrorLines.Clear()
        }
    }
    finally {
        $script:OpeningLog = $false
    }
}

function Ensure-RunLog {
    if ($script:RunSummaryLog -or $script:OpeningLog) {
        return
    }

    if ([string]::IsNullOrWhiteSpace($script:UtilityPath)) {
        return
    }

    try {
        Open-RunLogs -LogRoot (Join-Path $script:UtilityPath 'Logs')
    }
    catch {
        Write-Host "Unable to open the run log. $(ConvertTo-SingleLine $_.Exception.Message)" -ForegroundColor Yellow
    }
}

function Test-CsvContent {
    param([Parameter(Mandatory = $true)][string]$FilePath)

    $stream = $null
    $reader = $null
    try {
        $stream = [System.IO.File]::Open(
            $FilePath,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::Read
        )
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.UTF8Encoding]::new($false), $true)
        $line = $reader.ReadLine()
    }
    catch {
        return "The CSV could not be read. $(ConvertTo-SingleLine $_.Exception.Message)"
    }
    finally {
        if ($reader) { $reader.Dispose() }
        elseif ($stream) { $stream.Dispose() }
    }

    if ($null -eq $line -or [string]::IsNullOrWhiteSpace($line)) {
        return 'The CSV header row is empty.'
    }

    if ($line.IndexOf([char]0) -ge 0) {
        return 'The CSV file contains binary data and is not a text file.'
    }

    if ($line.Length -gt 200000) {
        return 'The CSV header row is longer than 200000 characters.'
    }

    $trimmed = $line.Trim()
    if ($trimmed -eq ',' -or $trimmed -eq ';' -or $trimmed -eq "`t") {
        return 'The CSV header row does not contain a column name.'
    }

    return $null
}

function Test-FreeSpaceForCopy {
    param(
        [Parameter(Mandatory = $true)][string]$Directory,
        [Parameter(Mandatory = $true)][long]$NeededBytes
    )

    try {
        $root = [System.IO.Path]::GetPathRoot($Directory)
        if ([string]::IsNullOrWhiteSpace($root)) {
            return $null
        }

        $drive = New-Object System.IO.DriveInfo $root
        if (-not $drive.IsReady) {
            return "The drive $root is not ready."
        }

        $required = $NeededBytes + 1MB
        if ($drive.AvailableFreeSpace -lt $required) {
            return "Not enough free space on $root. Need $required bytes, available $($drive.AvailableFreeSpace) bytes."
        }
    }
    catch {
        Write-RunLog -Level WARNING -Message "Free space could not be checked for $Directory. $($_.Exception.Message)"
    }

    return $null
}

function Get-FailureCategory {
    param([AllowNull()][string]$Text)

    $value = ''
    if ($Text) { $value = $Text.ToLowerInvariant() }

    if ($value -match 'timed out|timeout') { return 'TIMEOUT' }
    if ($value -match 'invalid client|unauthorized|http 401|authentication') { return 'AUTHENTICATION' }
    if ($value -match 'forbidden|access denied|http 403') { return 'AUTHORIZATION' }
    if ($value -match 'not found|http 404') { return 'NOT_FOUND' }
    if ($value -match 'bad request|http 400|http 409') { return 'BAD_REQUEST' }
    if ($value -match 'http 429|too many') { return 'RATE_LIMIT' }
    if ($value -match 'http 500|http 502|http 503|http 504') { return 'SERVER' }
    if ($value -match 'connection refused|unable to connect|could not resolve|no such host|network') { return 'NETWORK' }
    return 'JAVA_EXIT'
}

function Stop-UploadRun {
    param(
        [Parameter(Mandatory = $true)][int]$Code,
        [string]$Message,
        [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR')]
        [string]$Level = 'ERROR'
    )

    $script:RequestedExitCode = $Code
    $script:StopRequested = $true
    Ensure-RunLog

    if ($Message) {
        Write-RunLog -Level $Level -Message $Message
    }

    if ($script:RunSummaryLog -and -not $script:SummaryFinalized) {
        try {
            Write-Utf8Lines -Path $script:RunSummaryLog -Append -Lines @(
                ''
                "Run aborted. Exit code: $Code"
            )
            $script:SummaryFinalized = $true
        }
        catch {
            Write-Host "Unable to append abort footer: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    throw "UPLOAD_STOP:$Code"
}

function Test-UploadStop {
    param($ErrorRecord)

    if ($script:StopRequested) {
        return $true
    }

    $message = [string]$ErrorRecord.Exception.Message
    return $message -like '*UPLOAD_STOP:*'
}

function Format-CommandLineArgument {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Argument)

    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append('"')
    $pendingSlashes = 0

    foreach ($char in $Argument.ToCharArray()) {
        if ($char -eq '\') {
            $pendingSlashes++
            continue
        }

        if ($char -eq '"') {
            [void]$builder.Append('\' * (($pendingSlashes * 2) + 1))
            [void]$builder.Append('"')
            $pendingSlashes = 0
            continue
        }

        if ($pendingSlashes -gt 0) {
            [void]$builder.Append('\' * $pendingSlashes)
            $pendingSlashes = 0
        }

        [void]$builder.Append($char)
    }

    if ($pendingSlashes -gt 0) {
        [void]$builder.Append('\' * ($pendingSlashes * 2))
    }

    [void]$builder.Append('"')
    return $builder.ToString()
}

function ConvertTo-SafeFileName {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return 'unnamed'
    }

    $safe = ($Value.Trim() -replace '[^A-Za-z0-9_-]', '_')
    if ([string]::IsNullOrWhiteSpace($safe)) {
        return 'unnamed'
    }

    if ($safe.Length -gt 80) {
        $safe = $safe.Substring(0, 80)
    }

    return $safe
}

function ConvertTo-AggregationMode {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return 'account'
    }

    switch ($Value.Trim().ToLowerInvariant()) {
        'account' { return 'account' }
        'accounts' { return 'account' }
        'group' { return 'group' }
        'groups' { return 'group' }
        'both' { return 'both' }
        'all' { return 'both' }
        default { return $null }
    }
}

function Resolve-CsvAggregation {
    param([Parameter(Mandatory = $true)][System.IO.FileInfo]$CsvFile)

    if ($CsvFile.Name -match '(?i)^(?<source>.+)-group\.csv$') {
        return [PSCustomObject]@{
            SourceName      = $Matches['source'].Trim()
            AggregationType = 'group'
        }
    }

    return [PSCustomObject]@{
        SourceName      = $CsvFile.BaseName.Trim()
        AggregationType = 'account'
    }
}

function Get-SourceStateFileName {
    param(
        [Parameter(Mandatory = $true)][string]$SourceName,
        [string]$AggregationType = 'account'
    )

    $stateKey = $SourceName
    if ($AggregationType -eq 'group') {
        $stateKey = "group:$SourceName"
    }

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($stateKey)
        $hashBytes = $sha.ComputeHash($bytes)
        $hash = ([System.BitConverter]::ToString($hashBytes)).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }

    return "$hash.state"
}

function Get-RunStamp {
    return (Get-Date -Format 'yyyyMMdd_HH\hmm\Mss\s_fff')
}

function Resolve-BoundOrDefault {
    param(
        [Parameter(Mandatory = $true)][string]$ParameterName,
        [string[]]$EnvironmentNames,
        $ConfigValue,
        $DefaultValue
    )

    if ($script:BoundParameters.ContainsKey($ParameterName)) {
        return $script:BoundParameters[$ParameterName]
    }

    foreach ($name in $EnvironmentNames) {
        $envValue = [Environment]::GetEnvironmentVariable($name)
        if (-not [string]::IsNullOrWhiteSpace($envValue)) {
            return $envValue.Trim()
        }
    }

    if ($null -ne $ConfigValue -and "$ConfigValue".Trim() -ne '') {
        return $ConfigValue
    }

    return $DefaultValue
}

function ConvertTo-BooleanSetting {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        $Value
    )

    if ($Value -is [bool]) {
        return $Value
    }

    if ($Value -is [switch]) {
        return [bool]$Value
    }

    $text = "$Value".Trim()
    switch -Regex ($text) {
        '^(1|true|yes|y)$' { return $true }
        '^(0|false|no|n)$' { return $false }
        default {
            Stop-UploadRun -Code 1 -Message "Setting '$Name' must be true or false. Value: $text"
        }
    }
}

function ConvertTo-IntegerSetting {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        $Value,
        [int]$Minimum,
        [int]$Maximum
    )

    $number = 0
    if ($Value -is [int] -or $Value -is [long]) {
        $number = [int]$Value
    }
    elseif (-not [int]::TryParse("$Value".Trim(), [ref]$number)) {
        Stop-UploadRun -Code 1 -Message "Setting '$Name' must be an integer. Value: $Value"
    }

    if ($number -lt $Minimum -or $number -gt $Maximum) {
        Stop-UploadRun -Code 1 -Message "Setting '$Name' must be between $Minimum and $Maximum. Value: $number"
    }

    return $number
}

function Test-SafeFolderName {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$SettingName
    )

    if ([string]::IsNullOrWhiteSpace($Name) -or $Name -match '[\\/:*?"<>|]' -or $Name.Contains('..')) {
        Stop-UploadRun -Code 1 -Message "Setting '$SettingName' must be a single folder name. Value: $Name"
    }
}

function Read-UploadConfig {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $null
    }

    try {
        $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) {
            Stop-UploadRun -Code 1 -Message "Config file is empty: $Path"
        }

        $config = $raw | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        if (Test-UploadStop $_) { throw }
        Stop-UploadRun -Code 1 -Message "Config file could not be read: $Path. $($_.Exception.Message)"
    }

    $forbidden = @(
        'clientid', 'clientsecret', 'client_id', 'client_secret',
        'accesstoken', 'access_token', 'password', 'secret'
    )

    foreach ($property in @($config.PSObject.Properties)) {
        if ($forbidden -contains $property.Name.ToLowerInvariant()) {
            Stop-UploadRun -Code 1 -Message "Config file must not contain secrets. Remove '$($property.Name)' from $Path and store credentials in environment variables or with -SaveCredential."
        }
    }

    return $config
}

function Get-ConfigValue {
    param(
        $Config,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if ($null -eq $Config) {
        return $null
    }

    $property = $Config.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function Resolve-UtilityPath {
    param([string]$RequestedPath)

    $candidate = $RequestedPath
    if ([string]::IsNullOrWhiteSpace($candidate)) {
        $candidate = [Environment]::GetEnvironmentVariable('SAILPOINT_UTILITY_DIR')
    }

    if ([string]::IsNullOrWhiteSpace($candidate) -and -not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $candidate = $PSScriptRoot
    }

    if ([string]::IsNullOrWhiteSpace($candidate)) {
        $candidate = 'C:\SailPointUtility\FileUploadUtility'
    }

    try {
        $fullPath = [System.IO.Path]::GetFullPath($candidate.Trim())
    }
    catch {
        Stop-UploadRun -Code 1 -Message "UtilityDir is not a valid path. $($_.Exception.Message)"
    }

    $root = [System.IO.Path]::GetPathRoot($fullPath)
    if ($fullPath.TrimEnd('\') -eq $root.TrimEnd('\')) {
        Stop-UploadRun -Code 1 -Message "UtilityDir must not be a drive root: $fullPath"
    }

    return $fullPath
}

function Resolve-AnyDrivePath {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$RequestedPath,
        [Parameter(Mandatory = $true)][string]$BasePath,
        [Parameter(Mandatory = $true)][string]$SettingName
    )

    $candidate = $RequestedPath.Trim().Trim('"')
    if ([string]::IsNullOrWhiteSpace($candidate)) {
        Stop-UploadRun -Code 1 -Message "$SettingName is empty."
    }

    if ($candidate -match '[*?"<>|]') {
        Stop-UploadRun -Code 1 -Message "$SettingName contains invalid path characters. Value: $candidate"
    }

    $rooted = $false
    try {
        $rooted = [System.IO.Path]::IsPathRooted($candidate)
    }
    catch {
        Stop-UploadRun -Code 1 -Message "$SettingName is not a valid path. Value: $candidate"
    }

    if (-not $rooted) {
        $candidate = Join-Path $BasePath $candidate
    }

    try {
        $fullPath = [System.IO.Path]::GetFullPath($candidate)
    }
    catch {
        Stop-UploadRun -Code 1 -Message "$SettingName is not a valid path. $($_.Exception.Message)"
    }

    $root = [System.IO.Path]::GetPathRoot($fullPath)
    if ([string]::IsNullOrWhiteSpace($root) -or $fullPath.TrimEnd('\', '/') -eq $root.TrimEnd('\', '/')) {
        Stop-UploadRun -Code 1 -Message "$SettingName must be a folder path, not a drive root: $fullPath"
    }

    return $fullPath
}

function Set-OwnerOnlyAcl {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [switch]$Directory
    )

    try {
        $acl = New-Object System.Security.AccessControl.FileSecurity
        if ($Directory) {
            $acl = New-Object System.Security.AccessControl.DirectorySecurity
        }

        $acl.SetAccessRuleProtection($true, $false)
        $identities = @(
            [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            'NT AUTHORITY\SYSTEM'
        )

        $inherit = [System.Security.AccessControl.InheritanceFlags]::None
        $propagate = [System.Security.AccessControl.PropagationFlags]::None
        if ($Directory) {
            $inherit = [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit'
        }

        foreach ($identity in $identities) {
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                $identity,
                [System.Security.AccessControl.FileSystemRights]::FullControl,
                $inherit,
                $propagate,
                [System.Security.AccessControl.AccessControlType]::Allow
            )
            $acl.AddAccessRule($rule)
        }

        Set-Acl -LiteralPath $Path -AclObject $acl
    }
    catch {
        Write-RunLog -Level WARNING -Message "Permissions could not be restricted on $Path. $($_.Exception.Message)"
    }
}

function Unprotect-SecretText {
    param([Parameter(Mandatory = $true)][string]$Path)

    $secure = Import-Clixml -LiteralPath $Path
    if ($secure -isnot [System.Security.SecureString]) {
        throw "Encrypted secret file is invalid: $Path"
    }

    $pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer)
    }
}

function Save-SailPointCredential {
    param([Parameter(Mandatory = $true)][string]$SecretsDir)

    if (-not [Environment]::UserInteractive) {
        Stop-UploadRun -Code 1 -Message '-SaveCredential requires an interactive session. Run it as the scheduled-task account in a console window.'
    }

    try {
        if ([Console]::IsInputRedirected) {
            Stop-UploadRun -Code 1 -Message '-SaveCredential cannot read a redirected console.'
        }
    }
    catch {
        # Older hosts may not expose IsInputRedirected.
    }

    if (-not (Test-Path -LiteralPath $SecretsDir -PathType Container)) {
        New-Item -Path $SecretsDir -ItemType Directory -Force | Out-Null
    }

    Set-OwnerOnlyAcl -Path $SecretsDir -Directory

    $clientIdText = Read-Host -Prompt 'SailPoint client ID'
    $clientSecretSecure = Read-Host -Prompt 'SailPoint client secret' -AsSecureString

    if ([string]::IsNullOrWhiteSpace($clientIdText) -or $null -eq $clientSecretSecure -or $clientSecretSecure.Length -lt 1) {
        Stop-UploadRun -Code 1 -Message 'Client ID and client secret are both required.'
    }

    if ($clientIdText.Length -gt 256 -or $clientSecretSecure.Length -gt 512) {
        Stop-UploadRun -Code 1 -Message 'Client ID or client secret is longer than expected.'
    }

    $clientIdSecure = ConvertTo-SecureString -String $clientIdText.Trim() -AsPlainText -Force
    $clientIdPath = Join-Path $SecretsDir 'sail-client-id.xml'
    $clientSecretPath = Join-Path $SecretsDir 'sail-client-secret.xml'

    $clientIdSecure | Export-Clixml -LiteralPath $clientIdPath -Force
    $clientSecretSecure | Export-Clixml -LiteralPath $clientSecretPath -Force
    Set-OwnerOnlyAcl -Path $clientIdPath
    Set-OwnerOnlyAcl -Path $clientSecretPath

    $clientIdText = $null
    $clientIdSecure = $null
    $clientSecretSecure = $null

    Write-Host "Credentials saved for the current user under $SecretsDir" -ForegroundColor Green
    Write-Host 'The scheduled task must run as this same Windows user.' -ForegroundColor Yellow
}

function Get-SailPointCredential {
    param([Parameter(Mandatory = $true)][string]$SecretsDir)

    $sailId = [Environment]::GetEnvironmentVariable('SAIL_CLIENT_ID')
    $sailSecret = [Environment]::GetEnvironmentVariable('SAIL_CLIENT_SECRET')
    $legacyId = [Environment]::GetEnvironmentVariable('SAILPOINT_CLIENT_ID')
    $legacySecret = [Environment]::GetEnvironmentVariable('SAILPOINT_CLIENT_SECRET')

    $clientId = $null
    $clientSecret = $null

    if (-not [string]::IsNullOrWhiteSpace($sailId) -or -not [string]::IsNullOrWhiteSpace($sailSecret)) {
        if ([string]::IsNullOrWhiteSpace($sailId) -or [string]::IsNullOrWhiteSpace($sailSecret)) {
            Stop-UploadRun -Code 1 -Message 'Set both SAIL_CLIENT_ID and SAIL_CLIENT_SECRET. One of them is empty.'
        }

        $clientId = $sailId.Trim()
        $clientSecret = $sailSecret.Trim()
        $script:CredentialSource = 'SAIL_CLIENT_ID and SAIL_CLIENT_SECRET environment variables'
    }
    elseif (-not [string]::IsNullOrWhiteSpace($legacyId) -or -not [string]::IsNullOrWhiteSpace($legacySecret)) {
        if ([string]::IsNullOrWhiteSpace($legacyId) -or [string]::IsNullOrWhiteSpace($legacySecret)) {
            Stop-UploadRun -Code 1 -Message 'Set both SAILPOINT_CLIENT_ID and SAILPOINT_CLIENT_SECRET. One of them is empty.'
        }

        $clientId = $legacyId.Trim()
        $clientSecret = $legacySecret.Trim()
        $script:CredentialSource = 'SAILPOINT_CLIENT_ID and SAILPOINT_CLIENT_SECRET environment variables'
    }
    else {
        $clientIdPath = Join-Path $SecretsDir 'sail-client-id.xml'
        $clientSecretPath = Join-Path $SecretsDir 'sail-client-secret.xml'
        $idExists = Test-Path -LiteralPath $clientIdPath -PathType Leaf
        $secretExists = Test-Path -LiteralPath $clientSecretPath -PathType Leaf

        if ($idExists -or $secretExists) {
            if (-not $idExists -or -not $secretExists) {
                Stop-UploadRun -Code 1 -Message "Both DPAPI files are required in $SecretsDir (sail-client-id.xml and sail-client-secret.xml)."
            }

            try {
                $clientId = (Unprotect-SecretText -Path $clientIdPath).Trim()
                $clientSecret = (Unprotect-SecretText -Path $clientSecretPath).Trim()
                $script:CredentialSource = 'DPAPI files sail-client-id.xml and sail-client-secret.xml'
            }
            catch {
                if (Test-UploadStop $_) { throw }
                Stop-UploadRun -Code 1 -Message "Stored credentials could not be decrypted. Run -SaveCredential as the scheduled-task account. $($_.Exception.Message)"
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($clientId) -or [string]::IsNullOrWhiteSpace($clientSecret)) {
        Stop-UploadRun -Code 1 -Message 'SailPoint credentials were not found. Set SAIL_CLIENT_ID and SAIL_CLIENT_SECRET, or run -SaveCredential as the task account.'
    }

    Register-RedactedValue -Value $clientId
    Register-RedactedValue -Value $clientSecret
    $script:ClientId = $clientId
    $script:ClientSecret = $clientSecret
}

function Enter-ExecutionLock {
    param([Parameter(Mandatory = $true)][string]$Path)

    $script:LockPath = $Path

    try {
        $script:LockStream = [System.IO.File]::Open(
            $Path,
            [System.IO.FileMode]::OpenOrCreate,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None
        )
    }
    catch {
        $exception = $_.Exception
        while ($exception.InnerException) {
            $exception = $exception.InnerException
        }

        if ($exception -is [System.IO.IOException]) {
            Stop-UploadRun -Code 3 -Message "Another upload execution is already running. Lock file: $Path"
        }

        Stop-UploadRun -Code 1 -Message "Unable to create execution lock. $($exception.Message)"
    }

    $existingText = ''
    try {
        $script:LockStream.Position = 0
        $reader = New-Object System.IO.StreamReader(
            $script:LockStream,
            [System.Text.Encoding]::UTF8,
            $true,
            1024,
            $true
        )
        try {
            $existingText = $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }
    catch {
        $existingText = ''
    }

    $existingPid = $null
    $startedText = $null
    foreach ($line in @($existingText -split '\r?\n')) {
        if ($line -like 'PID=*') {
            $pidText = $line.Substring(4).Trim()
            $parsedPid = 0
            if ([int]::TryParse($pidText, [ref]$parsedPid) -and $parsedPid -gt 0) {
                $existingPid = $parsedPid
            }
        }
        elseif ($line -like 'Started=*') {
            $startedText = $line.Substring(8).Trim()
        }
    }

    if ($existingPid -and $existingPid -ne $PID) {
        $existingProcess = Get-Process -Id $existingPid -ErrorAction SilentlyContinue
        $stillRunning = $false
        if ($null -ne $existingProcess) {
            $stillRunning = $true
            if ($startedText) {
                try {
                    $lockStarted = [datetime]::ParseExact($startedText, 'yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
                    if ($existingProcess.StartTime -gt $lockStarted.AddSeconds(2)) {
                        $stillRunning = $false
                    }
                }
                catch {
                    $stillRunning = $true
                }
            }
        }

        if ($stillRunning) {
            if ($script:LockStream) {
                $script:LockStream.Dispose()
                $script:LockStream = $null
            }

            Stop-UploadRun -Code 3 -Message "Another upload execution is already running (PID $existingPid). Lock file: $Path"
        }
    }

    $lockLines = @(
        "Started=$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"
        "Computer=$env:COMPUTERNAME"
        "User=$env:USERNAME"
        "PID=$PID"
        "Version=$script:ScriptVersion"
    )
    $payload = [System.Text.Encoding]::UTF8.GetBytes(($lockLines -join "`r`n"))
    $script:LockStream.SetLength(0)
    $script:LockStream.Write($payload, 0, $payload.Length)
    $script:LockStream.Flush()
    $script:LockHeld = $true
}

function Exit-ExecutionLock {
    if ($script:LockStream) {
        try {
            $script:LockStream.Dispose()
        }
        catch {
            Write-RunLog -Level WARNING -Message 'Execution lock handle could not be released.'
        }

        $script:LockStream = $null
    }

    if ($script:LockHeld -and $script:LockPath -and (Test-Path -LiteralPath $script:LockPath -PathType Leaf)) {
        try {
            Remove-Item -LiteralPath $script:LockPath -Force -ErrorAction Stop
            $script:LockHeld = $false
        }
        catch {
            Write-RunLog -Level WARNING -Message "Execution lock could not be removed. $($_.Exception.Message) Lock file: $($script:LockPath)"
        }
    }
}

function Clear-UploadSecrets {
    $script:ClientId = $null
    $script:ClientSecret = $null
    $script:RedactedValues.Clear()
}

function Get-HttpStatusCode {
    param($ErrorRecord)

    $candidates = @(
        $ErrorRecord.Exception
        $ErrorRecord.Exception.InnerException
    )

    foreach ($candidate in $candidates) {
        if ($null -eq $candidate) {
            continue
        }

        $property = $candidate.PSObject.Properties['Response']
        if ($null -eq $property -or $null -eq $property.Value) {
            continue
        }

        try {
            return [int]$property.Value.StatusCode
        }
        catch {
            return $null
        }
    }

    return $null
}

function Invoke-SailPointWebRequest {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('GET', 'POST')][string]$Method,
        [Parameter(Mandatory = $true)][string]$Uri,
        [hashtable]$Headers,
        [string]$Body,
        [string]$ContentType,
        [int]$TimeoutSec,
        [int]$RetryCount
    )

    $attempt = 0
    while ($attempt -lt $RetryCount) {
        $attempt++
        $previousProgress = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        try {
            $request = @{
                Method          = $Method
                Uri             = $Uri
                TimeoutSec      = $TimeoutSec
                ErrorAction     = 'Stop'
            }

            if ($Headers) {
                $request.Headers = $Headers
            }

            if ($PSBoundParameters.ContainsKey('Body')) {
                $request.Body = $Body
            }

            if ($ContentType) {
                $request.ContentType = $ContentType
            }

            if ($PSVersionTable.PSVersion.Major -lt 6) {
                $request.UseBasicParsing = $true
            }

            return Invoke-WebRequest @request
        }
        catch {
            $status = Get-HttpStatusCode $_
            $retryable = ($null -eq $status) -or ($status -in 408, 429, 500, 502, 503, 504)
            if (-not $retryable -or $attempt -ge $RetryCount) {
                $safe = ConvertTo-SingleLine (Hide-SecretText $_.Exception.Message)
                $statusText = 'none'
                if ($null -ne $status) {
                    $statusText = [string]$status
                }

                throw "SailPoint request failed. HTTP status: $statusText. $safe"
            }

            $delaySeconds = [Math]::Min(30, [Math]::Pow(2, $attempt))
            $statusText = 'none'
            if ($null -ne $status) {
                $statusText = [string]$status
            }

            $uriPath = $Uri
            try { $uriPath = ([uri]$Uri).AbsolutePath } catch { }

            Write-RunLog -Level WARNING -Message "SailPoint request will be retried in $delaySeconds second(s). Attempt $attempt of $RetryCount. HTTP status: $statusText. Method=$Method Path=$uriPath"
            Start-Sleep -Seconds $delaySeconds
        }
        finally {
            $ProgressPreference = $previousProgress
        }
    }
}

function Get-SailPointAccessToken {
    param(
        [Parameter(Mandatory = $true)][string]$BaseUrl,
        [int]$RetryCount
    )

    $tokenUrl = "$BaseUrl/oauth/token"
    $body = 'grant_type=client_credentials&client_id={0}&client_secret={1}' -f @(
        [uri]::EscapeDataString($script:ClientId)
        [uri]::EscapeDataString($script:ClientSecret)
    )

    try {
        $response = Invoke-SailPointWebRequest -Method POST -Uri $tokenUrl -Body $body -ContentType 'application/x-www-form-urlencoded' -TimeoutSec 60 -RetryCount $RetryCount
    }
    catch {
        if (Test-UploadStop $_) { throw }
        throw "SailPoint authentication failed. $(ConvertTo-SingleLine (Hide-SecretText $_.Exception.Message))"
    }

    if ($null -eq $response -or [string]::IsNullOrWhiteSpace($response.Content)) {
        throw 'SailPoint authentication returned an empty response.'
    }

    try {
        $parsed = $response.Content | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw 'SailPoint authentication response was not valid JSON.'
    }

    $tokenProperty = $parsed.PSObject.Properties['access_token']
    if ($null -eq $tokenProperty -or [string]::IsNullOrWhiteSpace([string]$tokenProperty.Value)) {
        throw 'SailPoint authentication response did not contain an access token.'
    }

    $token = [string]$tokenProperty.Value
    Register-RedactedValue -Value $token
    return $token
}

function ConvertTo-DictionarySource {
    param($RawSource)

    if ($null -eq $RawSource) {
        return $null
    }

    $sourceName = ''
    $sourceId = ''

    if ($RawSource -is [System.Collections.IDictionary]) {
        foreach ($key in @($RawSource.Keys)) {
            $keyText = [string]$key
            if ($keyText -eq 'name') {
                $sourceName = [string]$RawSource[$key]
            }
            elseif ($keyText -eq 'id') {
                $sourceId = [string]$RawSource[$key]
            }
        }
    }
    else {
        $nameProperty = $RawSource.PSObject.Properties['name']
        $idProperty = $RawSource.PSObject.Properties['id']
        if ($nameProperty) {
            $sourceName = [string]$nameProperty.Value
        }
        if ($idProperty) {
            $sourceId = [string]$idProperty.Value
        }
    }

    $sourceName = $sourceName.Trim()
    $sourceId = $sourceId.Trim()
    if (-not $sourceName -or -not $sourceId) {
        return $null
    }

    return [PSCustomObject]@{
        name = $sourceName
        id   = $sourceId
    }
}

function ConvertFrom-SourceJson {
    param([AllowNull()][string]$JsonText)

    if ([string]::IsNullOrWhiteSpace($JsonText)) {
        return @()
    }

    if ($JsonText.Length -gt 67108864) {
        throw 'The Sources API response exceeds the 64 MB parse limit.'
    }

    $parsed = $null
    if ($PSVersionTable.PSVersion.Major -ge 6) {
        $parsed = $JsonText | ConvertFrom-Json -AsHashtable -Depth 100 -ErrorAction Stop
    }
    else {
        try {
            Add-Type -AssemblyName 'System.Web.Extensions' -ErrorAction Stop
            $serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
            $serializer.MaxJsonLength = 67108864
            $serializer.RecursionLimit = 100
            $parsed = $serializer.DeserializeObject($JsonText)
        }
        catch {
            Write-Verbose "JavaScriptSerializer was unavailable. Falling back to ConvertFrom-Json. $($_.Exception.Message)"
            $parsed = $JsonText | ConvertFrom-Json -ErrorAction Stop
        }
    }

    if ($null -eq $parsed) {
        return @()
    }

    $rawSources = @()
    if ($parsed -is [System.Collections.IDictionary]) {
        $keys = @($parsed.Keys | ForEach-Object { [string]$_ })
        if ($keys -contains 'items') {
            $rawSources = @($parsed['items'])
        }
        elseif ($keys -contains 'results') {
            $rawSources = @($parsed['results'])
        }
        elseif ($keys -contains 'data') {
            $rawSources = @($parsed['data'])
        }
        elseif ($keys -contains 'content') {
            $rawSources = @($parsed['content'])
        }
        elseif ($keys -contains 'name' -and $keys -contains 'id') {
            $rawSources = @($parsed)
        }
        else {
            throw "Unexpected Sources API response structure. Available keys: $($keys -join ', ')"
        }
    }
    elseif ($parsed -is [pscustomobject]) {
        $keys = @($parsed.PSObject.Properties.Name)
        if ($keys -contains 'items') {
            $rawSources = @($parsed.items)
        }
        elseif ($keys -contains 'results') {
            $rawSources = @($parsed.results)
        }
        elseif ($keys -contains 'data') {
            $rawSources = @($parsed.data)
        }
        elseif ($keys -contains 'content') {
            $rawSources = @($parsed.content)
        }
        elseif ($keys -contains 'name' -and $keys -contains 'id') {
            $rawSources = @($parsed)
        }
        else {
            throw "Unexpected Sources API response structure. Available keys: $($keys -join ', ')"
        }
    }
    elseif ($parsed -is [System.Collections.IEnumerable] -and -not ($parsed -is [string])) {
        $rawSources = @($parsed)
    }
    else {
        throw "Unexpected Sources API response type: $($parsed.GetType().FullName)"
    }

    $sources = New-Object System.Collections.Generic.List[object]
    foreach ($rawSource in $rawSources) {
        $simplified = ConvertTo-DictionarySource -RawSource $rawSource
        if ($null -ne $simplified) {
            $sources.Add($simplified)
        }
    }

    return @($sources.ToArray())
}

function Get-AllSailPointSources {
    param(
        [Parameter(Mandatory = $true)][string]$BaseUrl,
        [Parameter(Mandatory = $true)][string]$AccessToken,
        [int]$PageSize,
        [int]$RetryCount
    )

    $headers = @{
        Authorization = "Bearer $AccessToken"
        Accept        = 'application/json'
    }

    $allSources = New-Object System.Collections.Generic.List[object]
    $seenIds = @{}
    $offset = 0
    $pageGuard = 0

    do {
        $pageGuard++
        if ($pageGuard -gt 400) {
            throw 'Source paging stopped after 400 pages. The Sources API did not reach the last page.'
        }

        $sourcesUrl = '{0}/v3/sources?limit={1}&offset={2}' -f $BaseUrl, $PageSize, $offset
        $requestId = [guid]::NewGuid().ToString('n')
        $headers['X-Request-Id'] = $requestId

        try {
            $response = Invoke-SailPointWebRequest -Method GET -Uri $sourcesUrl -Headers $headers -TimeoutSec 120 -RetryCount $RetryCount
        }
        catch {
            if (Test-UploadStop $_) { throw }
            throw "Unable to retrieve SailPoint sources. RequestId=$requestId. $(ConvertTo-SingleLine (Hide-SecretText $_.Exception.Message))"
        }

        $page = @(ConvertFrom-SourceJson -JsonText $response.Content)
        Write-Verbose "Source page offset=$offset count=$($page.Count) requestId=$requestId"

        foreach ($source in $page) {
            if ($seenIds.ContainsKey($source.id)) {
                throw "Source paging returned duplicate id $($source.id). Offset paging did not advance."
            }

            $seenIds[$source.id] = $true
            $allSources.Add($source)
        }

        $offset += $PageSize
    } while ($page.Count -ge $PageSize)

    if ($allSources.Count -eq 0) {
        throw 'The Sources API returned no valid source records containing both name and id.'
    }

    return @($allSources.ToArray())
}

function New-SourceIndex {
    param(
        [object[]]$Sources,
        [bool]$CaseSensitive
    )

    $comparer = [StringComparer]::OrdinalIgnoreCase
    if ($CaseSensitive) {
        $comparer = [StringComparer]::Ordinal
    }

    $index = New-Object System.Collections.Hashtable $comparer
    foreach ($source in $Sources) {
        if ($null -eq $source) {
            continue
        }

        if (-not $index.ContainsKey($source.name)) {
            $index[$source.name] = New-Object System.Collections.Generic.List[object]
        }

        $index[$source.name].Add($source)
    }

    return $index
}

function Read-StateMap {
    param([string]$Path)

    $map = @{}
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $map
    }

    try {
        $lines = @(Get-Content -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop)
    }
    catch {
        Write-RunLog -Level WARNING -Message "State file could not be read and will be ignored. File=$Path Reason=$($_.Exception.Message)"
        return $map
    }

    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        $separator = $line.IndexOf('=')
        if ($separator -lt 1) {
            continue
        }

        $key = $line.Substring(0, $separator).Trim()
        $value = $line.Substring($separator + 1).Trim()
        $map[$key] = $value
    }

    return $map
}

function Write-StateFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$SourceName,
        [string]$SourceId,
        [string]$OriginalFile,
        [string]$Checksum,
        [long]$FileSize,
        [string]$LastWriteTimeUtc,
        [string]$SuccessfulUploadTime,
        [string]$ObjectType = 'account'
    )

    Write-Utf8Lines -Path $Path -Lines @(
        'Version=3'
        "SourceName=$SourceName"
        "ObjectType=$ObjectType"
        "SourceId=$SourceId"
        "OriginalFile=$OriginalFile"
        "Checksum=$Checksum"
        "FileSize=$FileSize"
        "LastWriteTimeUtc=$LastWriteTimeUtc"
        "SuccessfulUploadTime=$SuccessfulUploadTime"
    )
}

function Test-FileStable {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [int]$WaitSeconds
    )

    try {
        $first = Get-Item -LiteralPath $FilePath -ErrorAction Stop
        $firstLength = $first.Length
        $firstWrite = $first.LastWriteTimeUtc
        if ($WaitSeconds -gt 0) {
            Start-Sleep -Seconds $WaitSeconds
        }

        $second = Get-Item -LiteralPath $FilePath -ErrorAction Stop
        return ($firstLength -eq $second.Length -and $firstWrite -eq $second.LastWriteTimeUtc)
    }
    catch {
        return $false
    }
}

function Test-FileReadable {
    param([Parameter(Mandatory = $true)][string]$FilePath)

    $stream = $null
    try {
        $stream = [System.IO.File]::Open(
            $FilePath,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::None
        )
        return $true
    }
    catch {
        return $false
    }
    finally {
        if ($null -ne $stream) {
            $stream.Dispose()
        }
    }
}

function Get-FileChecksum {
    param([Parameter(Mandatory = $true)][string]$FilePath)

    $hash = Get-FileHash -LiteralPath $FilePath -Algorithm SHA256 -ErrorAction Stop
    return $hash.Hash.ToUpperInvariant()
}

function Get-FailureReason {
    param(
        [string[]]$OutputLines,
        [int]$ExitCode
    )

    $patterns = @(
        'unauthorized', 'forbidden', 'authentication', 'invalid client', 'access denied',
        'bad request', 'failed', 'failure', 'exception', 'error', 'timed out', 'timeout',
        'not found', 'connection refused', 'unable to', 'could not',
        'http 400', 'http 401', 'http 403', 'http 404', 'http 409', 'http 429',
        'http 500', 'http 502', 'http 503', 'http 504'
    )

    $reasonLines = New-Object System.Collections.Generic.List[string]
    foreach ($line in @($OutputLines)) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        foreach ($pattern in $patterns) {
            if ($line -match [regex]::Escape($pattern) -or $line.ToLowerInvariant().Contains($pattern)) {
                $reasonLines.Add((ConvertTo-SingleLine (Hide-SecretText $line)))
                break
            }
        }
    }

    $detail = $null
    if ($reasonLines.Count -gt 0) {
        $detail = $reasonLines[$reasonLines.Count - 1]
    }
    else {
        for ($i = $OutputLines.Count - 1; $i -ge 0; $i--) {
            if (-not [string]::IsNullOrWhiteSpace($OutputLines[$i])) {
                $detail = ConvertTo-SingleLine (Hide-SecretText $OutputLines[$i])
                break
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($detail)) {
        $detail = "Java utility returned exit code $ExitCode without a detailed error."
    }

    $category = Get-FailureCategory -Text $detail
    return "$category | $detail"
}

function Get-JavaMajorVersion {
    param([Parameter(Mandatory = $true)][string]$JavaExecutable)

    # java -version writes to stderr. With ErrorActionPreference Stop, PowerShell
    # turns that banner into a terminating error. Read both streams from the
    # process directly so the version text is data, not a script failure.
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $JavaExecutable
    $startInfo.Arguments = '-version'
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true

    $process = $null
    $stdout = ''
    $stderr = ''
    try {
        $process = [System.Diagnostics.Process]::Start($startInfo)
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit(15000)) {
            try { $process.Kill() } catch { }
            throw 'Java version check timed out.'
        }

        [void]$process.WaitForExit()
        if ($stdoutTask.Wait(5000) -and $stdoutTask.Status -eq [System.Threading.Tasks.TaskStatus]::RanToCompletion) {
            $stdout = [string]$stdoutTask.Result
        }
        if ($stderrTask.Wait(5000) -and $stderrTask.Status -eq [System.Threading.Tasks.TaskStatus]::RanToCompletion) {
            $stderr = [string]$stderrTask.Result
        }
    }
    finally {
        if ($process) {
            $process.Dispose()
        }
    }

    $versionText = ConvertTo-SingleLine "$stdout $stderr"
    $major = $null
    if ($versionText -match 'version "(?<major>\d+)\.') {
        $major = [int]$Matches.major
        if ($major -eq 1 -and $versionText -match 'version "1\.(?<legacy>\d+)\.') {
            $major = [int]$Matches.legacy
        }
    }

    return [PSCustomObject]@{
        Major = $major
        Text  = $versionText
    }
}

function Invoke-JavaUpload {
    param(
        [Parameter(Mandatory = $true)][string]$JavaExecutable,
        [Parameter(Mandatory = $true)][string]$JarPath,
        [Parameter(Mandatory = $true)][string]$BaseUrl,
        [Parameter(Mandatory = $true)][string]$FilePath,
        [int]$TimeoutMinutes,
        [string]$ObjectType = 'account'
    )

    $arguments = @(
        (Format-CommandLineArgument '-jar')
        (Format-CommandLineArgument $JarPath)
        (Format-CommandLineArgument '--url')
        (Format-CommandLineArgument $BaseUrl)
        (Format-CommandLineArgument '--clientId')
        (Format-CommandLineArgument 'env')
        (Format-CommandLineArgument '--clientSecret')
        (Format-CommandLineArgument 'env')
        (Format-CommandLineArgument '--objectType')
        (Format-CommandLineArgument $ObjectType)
        (Format-CommandLineArgument '-f')
        (Format-CommandLineArgument $FilePath)
        (Format-CommandLineArgument '--verbose')
    ) -join ' '

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $JavaExecutable
    $startInfo.Arguments = $arguments
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.RedirectStandardInput = $true
    $startInfo.EnvironmentVariables['SAIL_CLIENT_ID'] = $script:ClientId
    $startInfo.EnvironmentVariables['SAIL_CLIENT_SECRET'] = $script:ClientSecret

    $process = $null
    $stdoutTask = $null
    $stderrTask = $null
    $timedOut = $false
    $exitCode = -1
    $javaPid = ''
    $chunks = New-Object System.Collections.Generic.List[string]
    try {
        try {
            $process = [System.Diagnostics.Process]::Start($startInfo)
        }
        catch {
            $chunks.Add("Java process could not be started. $(Format-ErrorDetail $_)")
            return [PSCustomObject]@{
                ExitCode = -1
                JavaPid  = ''
                Output   = @($chunks.ToArray())
            }
        }

        $script:ActiveJavaProcess = $process
        $javaPid = [string]$process.Id
        Write-RunLog -Level INFO -Message "Java process started. PID=$javaPid ObjectType=$ObjectType File=$FilePath"
        try {
            $process.StandardInput.Close()
        }
        catch {
            Write-RunLog -Level WARNING -Message "Java process stdin could not be closed. PID=$($process.Id) Reason=$($_.Exception.Message)"
        }
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()

        $timeoutMs = [int]($TimeoutMinutes * 60 * 1000)
        if (-not $process.WaitForExit($timeoutMs)) {
            $timedOut = $true
            Write-RunLog -Level ERROR -Message "Java upload timed out after $TimeoutMinutes minute(s). PID=$($process.Id) ObjectType=$ObjectType File=$FilePath"
            try { $process.Kill() } catch {
                Write-RunLog -Level WARNING -Message "Timed-out Java process could not be stopped. PID=$($process.Id) Reason=$($_.Exception.Message)"
            }
            [void]$process.WaitForExit()
        }
        else {
            [void]$process.WaitForExit()
            $exitCode = $process.ExitCode
            Write-RunLog -Level INFO -Message "Java process finished. PID=$($process.Id) ExitCode=$exitCode ObjectType=$ObjectType"
        }
    }
    catch {
        $chunks.Add("Java upload failed. $(Format-ErrorDetail $_)")
        Write-RunLog -Level ERROR -Message "Java upload failed. $(Format-ErrorDetail $_)"
    }
    finally {
        foreach ($task in @($stdoutTask, $stderrTask)) {
            if ($null -eq $task) {
                continue
            }

            $completed = $false
            try {
                $completed = $task.Wait(10000)
            }
            catch {
                $completed = $false
            }

            if ($completed -and $task.Status -eq [System.Threading.Tasks.TaskStatus]::RanToCompletion -and $task.Result) {
                foreach ($line in @($task.Result -split '\r?\n')) {
                    $chunks.Add((Hide-SecretText $line))
                }
            }
            elseif (-not $completed) {
                $chunks.Add('Java output could not be read before the process was closed.')
                Write-RunLog -Level WARNING -Message 'Java output could not be read before the process was closed.'
            }
        }

        $script:ActiveJavaProcess = $null
        if ($process -and -not $process.HasExited) {
            try { $process.Kill() } catch {
                Write-RunLog -Level WARNING -Message "Java process was still running and could not be stopped. Reason=$($_.Exception.Message)"
            }
        }
        if ($process) {
            $process.Dispose()
        }
    }

    if ($timedOut) {
        $exitCode = -1
        $chunks.Add("Java upload timed out after $TimeoutMinutes minute(s).")
    }

    return [PSCustomObject]@{
        ExitCode = $exitCode
        JavaPid  = $javaPid
        Output   = @($chunks.ToArray())
    }
}

function Get-CsvHeaderNames {
    param([Parameter(Mandatory = $true)][string]$FilePath)

    $stream = $null
    $reader = $null
    try {
        $stream = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.UTF8Encoding]::new($false), $true)
        $line = $reader.ReadLine()
    }
    finally {
        if ($reader) { $reader.Dispose() }
        elseif ($stream) { $stream.Dispose() }
    }

    if ([string]::IsNullOrWhiteSpace($line)) {
        return @()
    }

    $names = New-Object System.Collections.Generic.List[string]
    foreach ($part in @($line.Split(','))) {
        $token = $part.Trim()
        if ($token.Length -ge 2 -and $token.StartsWith('"') -and $token.EndsWith('"')) {
            $token = $token.Substring(1, $token.Length - 2).Replace('""', '"').Trim()
        }
        if (-not [string]::IsNullOrWhiteSpace($token)) {
            $names.Add($token)
        }
    }

    return @($names.ToArray())
}

function Get-SourceSchemas {
    param(
        [Parameter(Mandatory = $true)][string]$BaseUrl,
        [Parameter(Mandatory = $true)][string]$SourceId,
        [int]$RetryCount
    )

    $token = Get-SailPointAccessToken -BaseUrl $BaseUrl -RetryCount $RetryCount
    try {
        $uri = "$BaseUrl/v3/sources/$SourceId/schemas"
        $headers = @{
            Authorization = "Bearer $token"
            Accept        = 'application/json'
        }
        $response = Invoke-SailPointWebRequest -Method GET -Uri $uri -Headers $headers -TimeoutSec 120 -RetryCount $RetryCount
        if ($null -eq $response -or [string]::IsNullOrWhiteSpace($response.Content)) {
            return @()
        }

        $parsed = $response.Content | ConvertFrom-Json -ErrorAction Stop
        return @($parsed)
    }
    finally {
        $token = $null
    }
}

function Get-SchemaAttributeNames {
    param($Schema)

    $raw = $null
    if ($Schema.PSObject.Properties['attributes']) {
        $raw = $Schema.attributes
    }
    elseif ($Schema.PSObject.Properties['attributeDefinitions']) {
        $raw = $Schema.attributeDefinitions
    }

    $names = New-Object System.Collections.Generic.List[string]
    foreach ($attribute in @($raw)) {
        if ($null -eq $attribute) {
            continue
        }

        $name = ''
        if ($attribute.PSObject.Properties['name']) {
            $name = [string]$attribute.name
        }
        elseif ($attribute.PSObject.Properties['attributeName']) {
            $name = [string]$attribute.attributeName
        }

        $name = $name.Trim()
        if ($name) {
            $names.Add($name)
        }
    }

    return @($names.ToArray())
}

function Get-GroupSchemaMismatch {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][string]$ObjectType,
        $Schemas
    )

    $available = New-Object System.Collections.Generic.List[string]
    $selected = $null
    foreach ($schema in @($Schemas)) {
        if ($null -eq $schema) {
            continue
        }

        $kind = ''
        foreach ($propertyName in @('nativeObjectType', 'objectType', 'name')) {
            $property = $schema.PSObject.Properties[$propertyName]
            if ($property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
                $kind = [string]$property.Value
                break
            }
        }

        $kind = $kind.Trim()
        if ($kind) {
            $available.Add($kind)
        }

        if ($kind -and $kind.Equals($ObjectType, [StringComparison]::OrdinalIgnoreCase)) {
            $selected = $schema
            break
        }
    }

    if ($null -eq $selected) {
        $known = 'none'
        if ($available.Count -gt 0) {
            $known = ($available -join ', ')
        }

        return "This source has no entitlement schema named '$ObjectType'. Schemas on the source: $known. Create that schema before uploading a group file."
    }

    $expected = @(Get-SchemaAttributeNames -Schema $selected)
    $actual = @(Get-CsvHeaderNames -FilePath $FilePath)
    if ($expected.Count -eq 0) {
        return $null
    }

    $expectedSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($name in $expected) { [void]$expectedSet.Add($name) }
    $actualSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($name in $actual) { [void]$actualSet.Add($name) }

    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($name in $expected) {
        if (-not $actualSet.Contains($name)) {
            $missing.Add($name)
        }
    }

    $unexpected = New-Object System.Collections.Generic.List[string]
    foreach ($name in $actual) {
        if (-not $expectedSet.Contains($name)) {
            $unexpected.Add($name)
        }
    }

    if ($missing.Count -eq 0 -and $unexpected.Count -eq 0) {
        return $null
    }

    return "The group CSV columns do not match entitlement schema '$ObjectType'. Expected: $($expected -join ', '). Missing: $($missing -join ', '). Unexpected: $($unexpected -join ', ')."
}

function Invoke-EntitlementUpload {
    param(
        [Parameter(Mandatory = $true)][string]$BaseUrl,
        [Parameter(Mandatory = $true)][string]$SourceId,
        [Parameter(Mandatory = $true)][string]$FilePath,
        [string]$ObjectType = 'group',
        [int]$TimeoutMinutes = 60,
        [int]$RetryCount = 3
    )

    $chunks = New-Object System.Collections.Generic.List[string]
    $chunks.Add("Submitting entitlement aggregation directly. SourceId=$SourceId ObjectType=$ObjectType File=$FilePath")
    $token = $null
    $requestStream = $null
    $response = $null
    $responseStream = $null
    $reader = $null

    try {
        $token = Get-SailPointAccessToken -BaseUrl $BaseUrl -RetryCount $RetryCount
        $boundary = '----SailPointGroup' + [guid]::NewGuid().ToString('n')
        $encoding = New-Object System.Text.UTF8Encoding $false
        $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
        $preamble = $encoding.GetBytes(
            "--$boundary`r`nContent-Disposition: form-data; name=`"objectType`"`r`n`r`n$ObjectType`r`n" +
            "--$boundary`r`nContent-Disposition: form-data; name=`"file`"; filename=`"file.csv`"`r`nContent-Type: text/csv`r`n`r`n"
        )
        $closing = $encoding.GetBytes("`r`n--$boundary--`r`n")
        $body = New-Object byte[] ($preamble.Length + $fileBytes.Length + $closing.Length)
        [System.Buffer]::BlockCopy($preamble, 0, $body, 0, $preamble.Length)
        [System.Buffer]::BlockCopy($fileBytes, 0, $body, $preamble.Length, $fileBytes.Length)
        [System.Buffer]::BlockCopy($closing, 0, $body, $preamble.Length + $fileBytes.Length, $closing.Length)

        $request = [System.Net.HttpWebRequest]::Create("$BaseUrl/beta/sources/$SourceId/load-entitlements")
        $request.Method = 'POST'
        $request.ContentType = "multipart/form-data; boundary=$boundary"
        $request.Accept = 'application/json'
        $request.Timeout = [int]($TimeoutMinutes * 60 * 1000)
        $request.ReadWriteTimeout = $request.Timeout
        $request.Headers.Add('Authorization', "Bearer $token")
        $request.ContentLength = $body.Length
        $requestStream = $request.GetRequestStream()
        $requestStream.Write($body, 0, $body.Length)
        $requestStream.Close()
        $requestStream = $null

        try {
            $response = $request.GetResponse()
        }
        catch {
            $candidate = $_.Exception
            while ($null -ne $candidate -and -not ($candidate -is [System.Net.WebException])) {
                $candidate = $candidate.InnerException
            }

            if ($candidate -and $candidate.Response) {
                $response = $candidate.Response
            }
            else {
                throw
            }
        }

        $statusCode = 0
        try { $statusCode = [int]$response.StatusCode } catch { $statusCode = 0 }
        $responseStream = $response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($responseStream)
        $responseText = $reader.ReadToEnd()
        $safeBody = ConvertTo-SingleLine (Hide-SecretText $responseText)
        if ([string]::IsNullOrWhiteSpace($safeBody)) {
            $safeBody = 'The response body was empty.'
        }

        if ($statusCode -ge 200 -and $statusCode -lt 300) {
            $chunks.Add("Entitlement aggregation accepted. HTTP status: $statusCode. $safeBody")
            return [PSCustomObject]@{
                ExitCode = 0
                JavaPid  = ''
                Output   = @($chunks.ToArray())
            }
        }

        $chunks.Add("Entitlement aggregation failed. HTTP status: $statusCode. $safeBody")
        return [PSCustomObject]@{
            ExitCode = $statusCode
            JavaPid  = ''
            Output   = @($chunks.ToArray())
        }
    }
    catch {
        $chunks.Add("Entitlement aggregation failed. $(Format-ErrorDetail $_)")
        return [PSCustomObject]@{
            ExitCode = -1
            JavaPid  = ''
            Output   = @($chunks.ToArray())
        }
    }
    finally {
        if ($reader) { $reader.Dispose() }
        if ($responseStream) { $responseStream.Dispose() }
        if ($requestStream) { $requestStream.Dispose() }
        if ($response) { $response.Dispose() }
        $token = $null
    }
}

function Write-AggregationLog {
    param(
        [Parameter(Mandatory = $true)][string]$LogFile,
        [string]$Status,
        [string]$Reason,
        [string]$OriginalFileName,
        [string]$OriginalFilePath,
        [string]$TemporaryFileName,
        [string]$TemporaryFilePath,
        [string]$SourceName,
        [string]$SourceId,
        [string]$Checksum,
        [long]$FileSize,
        [datetime]$StartTime,
        [datetime]$EndTime,
        [string]$ExitCode,
        [string]$JavaPid,
        [string]$TemporaryCleanupStatus,
        [string[]]$JavaOutput,
        [string]$TenantUrl,
        [string]$JarPath,
        [string]$MatchMode,
        [string]$AggregationType = 'account'
    )

    $durationSeconds = '{0:F2}' -f ($EndTime - $StartTime).TotalSeconds
    $sourceApiMatch = 'Not resolved'
    if ($SourceId) {
        $sourceApiMatch = $MatchMode
    }

    $javaLines = @($JavaOutput | ForEach-Object { Hide-SecretText $_ })
    if ($javaLines.Count -gt 400) {
        $omitted = $javaLines.Count - 400
        $javaLines = @($javaLines[0..199]) + @("... $omitted lines omitted ...") + @($javaLines[($javaLines.Count - 200)..($javaLines.Count - 1)])
    }

    if ($javaLines.Count -eq 0) {
        $javaLines = @('No output was returned by the Java utility.')
    }

    $content = @(
        '======================================================================'
        'SAILPOINT ISC DELIMITED SOURCE AGGREGATION LOG'
        '======================================================================'
        ''
        'FINAL RESULT'
        '----------------------------------------------------------------------'
        "Status                 : $Status"
        "Aggregation type       : $AggregationType"
        "Reason                 : $(Hide-SecretText $Reason)"
        "Java exit code         : $ExitCode"
        "Java PID               : $JavaPid"
        ''
        'SOURCE RESOLUTION'
        '----------------------------------------------------------------------'
        "Filename source name   : $SourceName"
        "Resolved Source ID     : $SourceId"
        "Source API match       : $sourceApiMatch"
        ''
        'ORIGINAL FILE'
        '----------------------------------------------------------------------'
        "Original filename      : $OriginalFileName"
        "Original path          : $OriginalFilePath"
        'Original retained      : YES'
        "File size              : $FileSize bytes"
        "SHA-256 checksum       : $Checksum"
        ''
        'TEMPORARY UPLOAD FILE'
        '----------------------------------------------------------------------'
        "Temporary filename     : $TemporaryFileName"
        "Temporary path         : $TemporaryFilePath"
        "Temporary cleanup      : $TemporaryCleanupStatus"
        ''
        'EXECUTION INFORMATION'
        '----------------------------------------------------------------------'
        "Computer               : $env:COMPUTERNAME"
        "Executed by            : $env:USERNAME"
        "PowerShell PID         : $PID"
        "Script version         : $script:ScriptVersion"
        "Tenant URL             : $TenantUrl"
        "JAR path               : $JarPath"
        "Start time             : $($StartTime.ToString('yyyy-MM-dd HH:mm:ss.fff'))"
        "End time               : $($EndTime.ToString('yyyy-MM-dd HH:mm:ss.fff'))"
        "Duration seconds       : $durationSeconds"
        ''
        'JAVA FILE UPLOAD UTILITY OUTPUT'
        '----------------------------------------------------------------------'
    ) + @($javaLines) + @(
        ''
        '======================================================================'
        'END OF AGGREGATION LOG'
        '======================================================================'
    )

    Write-Utf8Lines -Path $LogFile -Lines $content
}

function Invoke-OneCsvFile {
    param(
        [Parameter(Mandatory = $true)][System.IO.FileInfo]$CsvFile,
        [Parameter(Mandatory = $true)][int]$Position,
        [Parameter(Mandatory = $true)][int]$Total,
        $SourceIndex,
        $Settings
    )

    $result = [ordered]@{
        Status                   = 'FAILED'
        Reason                   = 'The file was not processed.'
        SourceName               = ''
        AggregationType          = ''
        SourceId                 = ''
        Checksum                 = ''
        TemporaryFileName        = ''
        TemporaryFilePath        = ''
        TemporaryCleanupStatus   = 'Not created'
        JavaExitCode             = ''
        JavaPid                  = ''
        JavaOutput               = @()
        StartTime                = Get-Date
        EndTime                  = Get-Date
        AggregationLog           = ''
        SummaryLogged            = $false
    }

    $plan = Resolve-CsvAggregation -CsvFile $CsvFile
    $sourceName = $plan.SourceName
    $aggregationType = $plan.AggregationType
    $safeSourceName = ConvertTo-SafeFileName -Value $sourceName
    $safeInputName = ConvertTo-SafeFileName -Value $CsvFile.BaseName
    $result.SourceName = $sourceName
    $result.AggregationType = $aggregationType
    $stamp = Get-RunStamp
    $aggregationLogName = "${safeInputName}_aggregation_log_${stamp}.txt"
    $result.AggregationLog = Join-Path $Settings.RunLogDir $aggregationLogName

    Write-Host ''
    Write-Host '============================================================' -ForegroundColor DarkGray
    Write-Host "Processing $Position of $Total" -ForegroundColor Cyan
    Write-Host "CSV file    : $($CsvFile.Name)" -ForegroundColor Gray
    Write-Host "Source name : $sourceName" -ForegroundColor Gray
    Write-Host "Object type : $aggregationType" -ForegroundColor Gray

    $temporaryPath = ''
    try {
        Write-RunLog -Level INFO -Message "FILE | $Position of $Total | ObjectType=$aggregationType | Source=$sourceName | File=$($CsvFile.Name) | Path=$($CsvFile.FullName)"

        if ($CsvFile.Name -like '~$*') {
            $result.Status = 'SKIPPED'
            $result.Reason = 'Temporary Office file ignored. It will not be uploaded.'
            throw 'FILE_OUTCOME'
        }

        if (-not $sourceName) {
            $result.Status = 'FAILED'
            $result.Reason = 'The CSV filename does not contain a source name.'
            throw 'FILE_OUTCOME'
        }

        if ($CsvFile.FullName.Length -gt 240) {
            $result.Status = 'FAILED'
            $result.Reason = "The CSV path is $($CsvFile.FullName.Length) characters. Windows cannot reliably open paths longer than 260 characters."
            throw 'FILE_OUTCOME'
        }

        if ($Settings.DuplicateTargets -and $Settings.DuplicateTargets.ContainsKey("$aggregationType|$sourceName")) {
            $result.Status = 'FAILED'
            $result.Reason = "More than one CSV in this run targets source '$sourceName' for $aggregationType. Files: $($Settings.DuplicateTargets["$aggregationType|$sourceName"]). No file was uploaded for this target."
            throw 'FILE_OUTCOME'
        }

        if ($CsvFile.Length -eq 0) {
            $result.Status = 'FAILED'
            $result.Reason = 'The CSV file is empty. A header row is required.'
            throw 'FILE_OUTCOME'
        }

        if (-not (Test-FileStable -FilePath $CsvFile.FullName -WaitSeconds $Settings.FileStabilityWaitSeconds)) {
            $result.Status = 'SKIPPED'
            $result.Reason = 'The CSV changed during the stability check or could not be read. The provisioning process may still be updating the file. It will be tried again on the next run.'
            throw 'FILE_OUTCOME'
        }

        if (-not (Test-FileReadable -FilePath $CsvFile.FullName)) {
            $result.Status = 'SKIPPED'
            $result.Reason = 'The CSV is locked or unavailable. The provisioning process may still be writing the file. It will be tried again on the next run.'
            throw 'FILE_OUTCOME'
        }

        $contentError = Test-CsvContent -FilePath $CsvFile.FullName
        if ($contentError) {
            $result.Status = 'FAILED'
            $result.Reason = $contentError
            throw 'FILE_OUTCOME'
        }

        $currentFile = Get-Item -LiteralPath $CsvFile.FullName -ErrorAction Stop
        $result.Checksum = Get-FileChecksum -FilePath $currentFile.FullName
        $statePath = Join-Path $Settings.StateDir (Get-SourceStateFileName -SourceName $sourceName -AggregationType $aggregationType)
        $legacyStatePath = Join-Path $Settings.StateDir "$safeSourceName.state"
        $stateMap = Read-StateMap -Path $statePath
        $stateSource = $statePath
        if ($aggregationType -eq 'account' -and -not $stateMap.ContainsKey('Checksum')) {
            $legacyMap = Read-StateMap -Path $legacyStatePath
            if ($legacyMap.ContainsKey('Checksum')) {
                $stateMap = $legacyMap
                $stateSource = $legacyStatePath
            }
        }

        $previousChecksum = ''
        if ($stateMap.ContainsKey('Checksum')) {
            $previousChecksum = $stateMap['Checksum'].ToUpperInvariant()
        }

        if ($previousChecksum -and $previousChecksum -eq $result.Checksum) {
            $result.Status = 'UNCHANGED'
            $result.Reason = 'The file checksum matches the last successfully aggregated version.'
            if ($stateSource -ne $statePath) {
                try {
                    Write-StateFile -Path $statePath -SourceName $sourceName -SourceId $stateMap['SourceId'] -OriginalFile $currentFile.Name -Checksum $result.Checksum -FileSize $currentFile.Length -LastWriteTimeUtc $currentFile.LastWriteTimeUtc.ToString('yyyy-MM-dd HH:mm:ss.fff') -SuccessfulUploadTime $stateMap['SuccessfulUploadTime'] -ObjectType $aggregationType
                }
                catch {
                    $script:CleanupWarningCount++
                    Write-RunLog -Level WARNING -Message "State migration failed | File=$($currentFile.Name) | Reason=$($_.Exception.Message)"
                }
            }
            throw 'FILE_OUTCOME'
        }

        if (-not $SourceIndex.ContainsKey($sourceName)) {
            $result.Status = 'FAILED'
            $result.Reason = "No SailPoint ISC source was found with the exact name '$sourceName'."
            throw 'FILE_OUTCOME'
        }

        $matched = $SourceIndex[$sourceName]
        if ($matched.Count -gt 1) {
            $result.Status = 'FAILED'
            $result.Reason = "Multiple SailPoint ISC sources were found with the exact name '$sourceName'. The upload destination cannot be selected safely."
            throw 'FILE_OUTCOME'
        }

        $sourceId = ([string]$matched[0].id).Trim()
        $result.SourceId = $sourceId
        Write-Host "Source ID   : $sourceId" -ForegroundColor Gray

        if (-not $sourceId -or $sourceId -notmatch '^[A-Fa-f0-9]{32}$') {
            $result.Status = 'FAILED'
            $result.Reason = "The resolved Source ID '$sourceId' is not a valid 32-character hexadecimal Source ID."
            throw 'FILE_OUTCOME'
        }

        if ($WhatIfPreference -or $Settings.WhatIf) {
            $result.Status = 'WHATIF'
            $result.Reason = 'WhatIf is set. The changed file was not uploaded.'
            throw 'FILE_OUTCOME'
        }

        if ($aggregationType -eq 'group') {
            $temporaryName = '{0}-{1}-group.csv' -f $sourceId, $safeSourceName
        }
        else {
            $temporaryName = '{0}-{1}.csv' -f $sourceId, $safeSourceName
        }
        $temporaryPath = Join-Path $Settings.ProcessingDir $temporaryName
        $result.TemporaryFileName = $temporaryName
        $result.TemporaryFilePath = $temporaryPath

        $spaceError = Test-FreeSpaceForCopy -Directory $Settings.ProcessingDir -NeededBytes $currentFile.Length
        if ($spaceError) {
            $result.Status = 'FAILED'
            $result.Reason = $spaceError
            throw 'FILE_OUTCOME'
        }

        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) {
            Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction Stop
        }

        Copy-Item -LiteralPath $currentFile.FullName -Destination $temporaryPath -Force -ErrorAction Stop
        $result.TemporaryCleanupStatus = 'Pending'
        $temporaryChecksum = Get-FileChecksum -FilePath $temporaryPath
        if ($temporaryChecksum -ne $result.Checksum) {
            throw 'Temporary file checksum does not match the original file checksum.'
        }

        if ($aggregationType -eq 'group') {
            try {
                $schemas = @(Get-SourceSchemas -BaseUrl $Settings.TenantUrl -SourceId $sourceId -RetryCount $Settings.ApiRetryCount)
                $schemaMismatch = Get-GroupSchemaMismatch -FilePath $temporaryPath -ObjectType $aggregationType -Schemas $schemas
            }
            catch {
                $result.Status = 'FAILED'
                $result.Reason = "The source entitlement schema could not be read. $(Format-ErrorDetail $_)"
                $result.SummaryLogged = $true
                Write-RunLog -Level ERROR -Message "FAILED | ObjectType=$aggregationType | Source=$sourceName | SourceId=$sourceId | File=$($currentFile.Name) | Reason=$($result.Reason)"
                throw 'FILE_OUTCOME'
            }

            if ($schemaMismatch) {
                $result.Status = 'FAILED'
                $result.Reason = $schemaMismatch
                $result.SummaryLogged = $true
                Write-RunLog -Level ERROR -Message "FAILED | ObjectType=$aggregationType | Source=$sourceName | SourceId=$sourceId | File=$($currentFile.Name) | Reason=$schemaMismatch"
                throw 'FILE_OUTCOME'
            }
        }

        Write-RunLog -Level INFO -Message "STARTED | ObjectType=$aggregationType | Source=$sourceName | SourceId=$sourceId | OriginalFile=$($currentFile.Name) | TemporaryFile=$temporaryName"

        if ($aggregationType -eq 'group') {
            $javaResult = Invoke-EntitlementUpload -BaseUrl $Settings.TenantUrl -SourceId $sourceId -FilePath $temporaryPath -ObjectType $aggregationType -TimeoutMinutes $Settings.JavaTimeoutMinutes -RetryCount $Settings.ApiRetryCount
        }
        else {
            $javaResult = Invoke-JavaUpload -JavaExecutable $Settings.JavaExecutable -JarPath $Settings.JarPath -BaseUrl $Settings.TenantUrl -FilePath $temporaryPath -TimeoutMinutes $Settings.JavaTimeoutMinutes -ObjectType $aggregationType
        }
        $result.JavaExitCode = [string]$javaResult.ExitCode
        $result.JavaPid = [string]$javaResult.JavaPid
        $result.JavaOutput = @($javaResult.Output)
        $result.EndTime = Get-Date

        if ($javaResult.ExitCode -eq 0) {
            $suspicious = @($result.JavaOutput | Where-Object { $_ -match '(?i)\b(exception|upload failed|aggregation failed)\b' } | Select-Object -First 1)
            if ($suspicious.Count -gt 0) {
                Write-RunLog -Level WARNING -Message "Java exit code was 0 and the output contains an error-like line. $(ConvertTo-SingleLine (Hide-SecretText $suspicious[0]))"
            }

            $uploadedAt = $result.EndTime.ToString('yyyy-MM-dd HH:mm:ss.fff')
            try {
                Write-StateFile -Path $statePath -SourceName $sourceName -SourceId $sourceId -OriginalFile $currentFile.Name -Checksum $result.Checksum -FileSize $currentFile.Length -LastWriteTimeUtc $currentFile.LastWriteTimeUtc.ToString('yyyy-MM-dd HH:mm:ss.fff') -SuccessfulUploadTime $uploadedAt -ObjectType $aggregationType
            }
            catch {
                $result.Status = 'FAILED'
                $result.Reason = "Aggregation upload completed, but the state file could not be saved. The file will be uploaded again on the next run. $(Format-ErrorDetail $_)"
                $result.SummaryLogged = $true
                Write-RunLog -Level ERROR -Message "FAILED | ObjectType=$aggregationType | Source=$sourceName | SourceId=$sourceId | File=$($currentFile.Name) | ExitCode=0 | Reason=$($result.Reason) | Log=$aggregationLogName"
                throw 'FILE_OUTCOME'
            }

            $result.Status = 'SUCCESS'
            $result.Reason = 'Aggregation upload completed successfully.'
            $result.SummaryLogged = $true
            Write-RunLog -Level SUCCESS -Message "SUCCESS | ObjectType=$aggregationType | Source=$sourceName | SourceId=$sourceId | File=$($currentFile.Name) | ExitCode=$($javaResult.ExitCode) | JavaPid=$($result.JavaPid) | Log=$aggregationLogName"
            throw 'FILE_OUTCOME'
        }

        $result.Status = 'FAILED'
        $result.Reason = Get-FailureReason -OutputLines $result.JavaOutput -ExitCode $javaResult.ExitCode
        $result.SummaryLogged = $true
        Write-RunLog -Level ERROR -Message "FAILED | ObjectType=$aggregationType | Source=$sourceName | SourceId=$sourceId | File=$($currentFile.Name) | ExitCode=$($javaResult.ExitCode) | Reason=$($result.Reason) | Log=$aggregationLogName"
    }
    catch {
        if (Test-UploadStop $_) {
            $result.Status = 'FAILED'
            $result.Reason = 'This file failed. The run will continue with the next file.'
            $result.SummaryLogged = $true
            Write-RunLog -Level ERROR -Message "FAILED | ObjectType=$aggregationType | Source=$sourceName | SourceId=$($result.SourceId) | File=$($CsvFile.Name) | Reason=$($result.Reason) | Log=$aggregationLogName"
        }
        elseif ($_.Exception.Message -eq 'FILE_OUTCOME') {
            # The file outcome is already stored on $result.
        }
        else {
            $result.EndTime = Get-Date
            $result.Status = 'FAILED'
            $result.Reason = Format-ErrorDetail $_
            if ($_.ScriptStackTrace) {
                $stack = ConvertTo-SingleLine (Hide-SecretText $_.ScriptStackTrace)
                Write-RunLog -Level ERROR -Message "STACK | File=$($CsvFile.Name) | $stack"
            }
            $result.JavaOutput = @($result.JavaOutput) + @("PowerShell error: $($result.Reason)")
            $result.SummaryLogged = $true
            Write-RunLog -Level ERROR -Message "FAILED | ObjectType=$aggregationType | Source=$sourceName | SourceId=$($result.SourceId) | File=$($CsvFile.Name) | Reason=$($result.Reason) | Log=$aggregationLogName"
        }
    }
    finally {
        if ($temporaryPath -and (Test-Path -LiteralPath $temporaryPath -PathType Leaf)) {
            try {
                Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction Stop
                $result.TemporaryCleanupStatus = 'SUCCESS'
            }
            catch {
                $result.TemporaryCleanupStatus = "FAILED: $($_.Exception.Message)"
                $script:CleanupWarningCount++
                Write-RunLog -Level WARNING -Message "Temporary file cleanup failed | File=$temporaryPath | Reason=$($_.Exception.Message)"
            }
        }
        elseif ($temporaryPath) {
            $result.TemporaryCleanupStatus = 'Already absent'
        }

        $result.EndTime = Get-Date
        try {
            Write-AggregationLog -LogFile $result.AggregationLog -Status $result.Status -Reason $result.Reason -OriginalFileName $CsvFile.Name -OriginalFilePath $CsvFile.FullName -TemporaryFileName $result.TemporaryFileName -TemporaryFilePath $result.TemporaryFilePath -SourceName $result.SourceName -SourceId $result.SourceId -Checksum $result.Checksum -FileSize $CsvFile.Length -StartTime $result.StartTime -EndTime $result.EndTime -ExitCode ([string]$result.JavaExitCode) -JavaPid ([string]$result.JavaPid) -TemporaryCleanupStatus $result.TemporaryCleanupStatus -JavaOutput $result.JavaOutput -TenantUrl $Settings.TenantUrl -JarPath $Settings.JarPath -MatchMode $Settings.MatchMode -AggregationType $aggregationType
        }
        catch {
            $script:CleanupWarningCount++
            Write-RunLog -Level WARNING -Message "Aggregation log creation failed | File=$($CsvFile.Name) | Log=$($result.AggregationLog) | Reason=$($_.Exception.Message)"
        }
    }

    return [PSCustomObject]$result
}

function Write-FinalSummary {
    param(
        $Settings,
        [datetime]$Started,
        [int]$FileCount,
        [int]$ExitCode
    )

    $ended = Get-Date
    $duration = '{0:F2}' -f ($ended - $Started).TotalSeconds
    $lines = @(
        ''
        '======================================================================'
        'FINAL RUN SUMMARY'
        '======================================================================'
        "Run started        : $($Started.ToString('yyyy-MM-dd HH:mm:ss.fff'))"
        "Run completed      : $($ended.ToString('yyyy-MM-dd HH:mm:ss.fff'))"
        "Total duration     : $duration seconds"
        "CSV files found    : $FileCount"
        "Successful uploads : $($script:SuccessCount)"
        "Failed uploads     : $($script:FailureCount)"
        "Skipped files      : $($script:SkippedCount)"
        "Unchanged files    : $($script:UnchangedCount)"
        "WhatIf uploads     : $($script:WhatIfCount)"
        "Cleanup warnings   : $($script:CleanupWarningCount)"
        "Exit code          : $ExitCode"
        "Summary log        : $($script:RunSummaryLog)"
        "Error log          : $($script:RunErrorLog)"
        'Original retained  : YES'
        'Archive used       : NO'
        '======================================================================'
    )

    Write-Utf8Lines -Path $script:RunSummaryLog -Lines $lines -Append
    $script:SummaryFinalized = $true

    Write-Host ''
    Write-Host '============================================================' -ForegroundColor DarkGray
    Write-Host 'Final Execution Summary' -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor DarkGray
    Write-Host "CSV files found    : $FileCount" -ForegroundColor White
    Write-Host "Successful uploads : $($script:SuccessCount)" -ForegroundColor Green
    Write-Host "Failed uploads     : $($script:FailureCount)" -ForegroundColor Red
    Write-Host "Skipped files      : $($script:SkippedCount)" -ForegroundColor Yellow
    Write-Host "Unchanged files    : $($script:UnchangedCount)" -ForegroundColor Gray
    Write-Host "WhatIf uploads     : $($script:WhatIfCount)" -ForegroundColor Gray
    Write-Host "Cleanup warnings   : $($script:CleanupWarningCount)" -ForegroundColor Yellow
    Write-Host "Summary log        : $($script:RunSummaryLog)" -ForegroundColor Gray
    Write-Host "Error log          : $($script:RunErrorLog)" -ForegroundColor Gray
    Write-Host '============================================================' -ForegroundColor DarkGray
}

function Remove-ExpiredRunLogs {
    param(
        [Parameter(Mandatory = $true)][string]$LogRoot,
        [int]$Keep
    )

    try {
        $oldFolders = @(
            Get-ChildItem -LiteralPath $LogRoot -Directory -Filter 'Run_*' -ErrorAction Stop |
                Sort-Object LastWriteTime -Descending |
                Select-Object -Skip $Keep
        )

        foreach ($folder in $oldFolders) {
            Remove-Item -LiteralPath $folder.FullName -Recurse -Force -ErrorAction Stop
            Write-Host "Removed old log folder: $($folder.Name)" -ForegroundColor Gray
        }
    }
    catch {
        $script:CleanupWarningCount++
        Write-RunLog -Level WARNING -Message "Log rotation warning: $($_.Exception.Message)"
    }
}

function Start-SailPointUpload {
    Enable-Tls12

    $utilityPath = Resolve-UtilityPath -RequestedPath $UtilityDir
    $script:UtilityPath = $utilityPath
    $configPathResolved = $ConfigPath
    if ([string]::IsNullOrWhiteSpace($configPathResolved)) {
        $configPathResolved = [Environment]::GetEnvironmentVariable('SAILPOINT_CONFIG_PATH')
    }
    if ([string]::IsNullOrWhiteSpace($configPathResolved)) {
        $configPathResolved = Join-Path $utilityPath 'upload-config.json'
    }
    else {
        try {
            $configPathResolved = [System.IO.Path]::GetFullPath($configPathResolved)
        }
        catch {
            Ensure-RunLog
            Stop-UploadRun -Code 1 -Message "Config path is not valid. $(Format-ErrorDetail $_)"
        }
    }

    $config = $null
    $configFailure = $null
    if (-not $SaveCredential) {
        try {
            $config = Read-UploadConfig -Path $configPathResolved
        }
        catch {
            $configFailure = $_
        }
    }

    $logFolder = 'Logs'
    $configuredLogFolder = [string](Get-ConfigValue $config 'logFolderName')
    if (-not [string]::IsNullOrWhiteSpace($configuredLogFolder)) {
        $logFolder = $configuredLogFolder.Trim()
    }
    if ([string]::IsNullOrWhiteSpace($logFolder) -or $logFolder -match '[\\/:*?"<>|]' -or $logFolder.Contains('..')) {
        $invalidLogFolder = $logFolder
        $logFolder = 'Logs'
    }
    else {
        $invalidLogFolder = $null
    }

    try {
        Open-RunLogs -LogRoot (Join-Path $utilityPath $logFolder)
    }
    catch {
        Write-Host "Unable to open the run log under $utilityPath\$logFolder. $(ConvertTo-SingleLine $_.Exception.Message)" -ForegroundColor Red
        if ($configFailure -and (Test-UploadStop $configFailure)) {
            throw $configFailure
        }
        throw
    }

    if ($script:JavaHomeNote) {
        Write-RunLog -Level WARNING -Message $script:JavaHomeNote
    }

    if ($invalidLogFolder) {
        Stop-UploadRun -Code 1 -Message "Setting 'logFolderName' must be a single folder name. Value: $invalidLogFolder"
    }

    if ($configFailure) {
        if (Test-UploadStop $configFailure) {
            if (-not $script:SummaryFinalized) {
                try {
                    Write-Utf8Lines -Path $script:RunSummaryLog -Append -Lines @('', "Run aborted. Exit code: $($script:RequestedExitCode)")
                    $script:SummaryFinalized = $true
                }
                catch { }
            }
            throw $configFailure
        }

        Stop-UploadRun -Code 1 -Message "Config file could not be read. $(Format-ErrorDetail $configFailure)"
    }

    if ($SaveCredential) {
        if (-not (Test-Path -LiteralPath $utilityPath -PathType Container)) {
            New-Item -Path $utilityPath -ItemType Directory -Force | Out-Null
        }

        try {
            Save-SailPointCredential -SecretsDir (Join-Path $utilityPath 'Secrets')
        }
        catch {
            if (Test-UploadStop $_) { throw }
            Stop-UploadRun -Code 1 -Message "Credentials could not be saved. $(Format-ErrorDetail $_)"
        }

        Write-RunLog -Level SUCCESS -Message 'Credentials saved for the current Windows user. The scheduled task must run as this same user.'
        Write-FinalSummary -Settings $null -Started $script:RunStarted -FileCount 0 -ExitCode 0
        return 0
    }

    $tenant = [string](Resolve-BoundOrDefault -ParameterName 'TenantUrl' -EnvironmentNames @('SAIL_BASE_URL', 'SAILPOINT_TENANT_URL') -ConfigValue (Get-ConfigValue $config 'tenantUrl') -DefaultValue '')
    $tenant = $tenant.Trim().TrimEnd('/')
    if ([string]::IsNullOrWhiteSpace($tenant)) {
        Stop-UploadRun -Code 1 -Message 'Tenant URL is empty. Pass -TenantUrl, set SAIL_BASE_URL, or set tenantUrl in upload-config.json.'
    }

    if ($tenant -notmatch '^https://[A-Za-z0-9.-]+$') {
        Stop-UploadRun -Code 1 -Message 'Tenant URL is invalid. Use only the plain HTTPS API URL, with no path, query, or credentials.'
    }

    $jar = [string](Resolve-BoundOrDefault -ParameterName 'JarName' -EnvironmentNames @('SAILPOINT_JAR_NAME') -ConfigValue (Get-ConfigValue $config 'jarName') -DefaultValue 'sailpoint-file-upload-utility-4.1.0.jar')
    $pageSize = ConvertTo-IntegerSetting -Name 'SourcesPageSize' -Value (Resolve-BoundOrDefault -ParameterName 'SourcesPageSize' -EnvironmentNames @('SAILPOINT_SOURCES_PAGE_SIZE') -ConfigValue (Get-ConfigValue $config 'sourcesPageSize') -DefaultValue 250) -Minimum 1 -Maximum 250
    $stabilitySeconds = ConvertTo-IntegerSetting -Name 'FileStabilityWaitSeconds' -Value (Resolve-BoundOrDefault -ParameterName 'FileStabilityWaitSeconds' -EnvironmentNames @('SAILPOINT_FILE_STABILITY_SECONDS') -ConfigValue (Get-ConfigValue $config 'fileStabilityWaitSeconds') -DefaultValue 5) -Minimum 0 -Maximum 3600
    $runsToKeepValue = ConvertTo-IntegerSetting -Name 'RunsToKeep' -Value (Resolve-BoundOrDefault -ParameterName 'RunsToKeep' -EnvironmentNames @('SAILPOINT_RUNS_TO_KEEP') -ConfigValue (Get-ConfigValue $config 'runsToKeep') -DefaultValue 90) -Minimum 1 -Maximum 3650
    $javaTimeout = ConvertTo-IntegerSetting -Name 'JavaTimeoutMinutes' -Value (Resolve-BoundOrDefault -ParameterName 'JavaTimeoutMinutes' -EnvironmentNames @('SAILPOINT_JAVA_TIMEOUT_MINUTES') -ConfigValue (Get-ConfigValue $config 'javaTimeoutMinutes') -DefaultValue 60) -Minimum 1 -Maximum 240
    $retryCount = ConvertTo-IntegerSetting -Name 'ApiRetryCount' -Value (Resolve-BoundOrDefault -ParameterName 'ApiRetryCount' -EnvironmentNames @('SAILPOINT_API_RETRY_COUNT') -ConfigValue (Get-ConfigValue $config 'apiRetryCount') -DefaultValue 3) -Minimum 1 -Maximum 5
    $aggregationMode = ConvertTo-AggregationMode -Value (Resolve-BoundOrDefault -ParameterName 'ObjectType' -EnvironmentNames @('SAILPOINT_OBJECT_TYPE') -ConfigValue (Get-ConfigValue $config 'objectType') -DefaultValue 'account')
    if (-not $aggregationMode) {
        $requestedObjectType = [string](Resolve-BoundOrDefault -ParameterName 'ObjectType' -EnvironmentNames @('SAILPOINT_OBJECT_TYPE') -ConfigValue (Get-ConfigValue $config 'objectType') -DefaultValue 'account')
        Stop-UploadRun -Code 1 -Message "Object type '$requestedObjectType' is invalid. Use account, group, or both."
    }

    $pause = ConvertTo-BooleanSetting -Name 'PauseAtEnd' -Value (Resolve-BoundOrDefault -ParameterName 'PauseAtEnd' -EnvironmentNames @('SAILPOINT_PAUSE_AT_END') -ConfigValue (Get-ConfigValue $config 'pauseAtEnd') -DefaultValue $false)
    $discovery = ConvertTo-BooleanSetting -Name 'EnableSourceDiscoveryLogging' -Value (Resolve-BoundOrDefault -ParameterName 'EnableSourceDiscoveryLogging' -EnvironmentNames @('SAILPOINT_LOG_SOURCE_DISCOVERY') -ConfigValue (Get-ConfigValue $config 'enableSourceDiscoveryLogging') -DefaultValue $false)
    $caseSensitive = ConvertTo-BooleanSetting -Name 'CaseSensitiveSourceMatch' -Value (Resolve-BoundOrDefault -ParameterName 'CaseSensitiveSourceMatch' -EnvironmentNames @('SAILPOINT_CASE_SENSITIVE_SOURCE_MATCH') -ConfigValue (Get-ConfigValue $config 'caseSensitiveSourceMatch') -DefaultValue $false)
    $whatIfEnabled = [bool]$WhatIf
    $script:PauseRequested = [bool]$pause

    $logFolder = [string](Get-ConfigValue $config 'logFolderName')
    if ([string]::IsNullOrWhiteSpace($logFolder)) { $logFolder = 'Logs' }
    $stateFolder = [string](Get-ConfigValue $config 'stateFolderName')
    if ([string]::IsNullOrWhiteSpace($stateFolder)) { $stateFolder = 'State' }

    Test-SafeFolderName -Name $logFolder -SettingName 'logFolderName'
    Test-SafeFolderName -Name $stateFolder -SettingName 'stateFolderName'
    Test-SafeFolderName -Name $jar -SettingName 'jarName'

    $sourceConfigured = Get-ConfigValue $config 'sourceDirectory'
    if ([string]::IsNullOrWhiteSpace([string]$sourceConfigured)) {
        $sourceConfigured = Get-ConfigValue $config 'inputDirectory'
    }
    if ([string]::IsNullOrWhiteSpace([string]$sourceConfigured)) {
        $sourceConfigured = Get-ConfigValue $config 'inputFolderName'
    }

    $temporaryConfigured = Get-ConfigValue $config 'temporaryDirectory'
    if ([string]::IsNullOrWhiteSpace([string]$temporaryConfigured)) {
        $temporaryConfigured = Get-ConfigValue $config 'processingDirectory'
    }
    if ([string]::IsNullOrWhiteSpace([string]$temporaryConfigured)) {
        $temporaryConfigured = Get-ConfigValue $config 'processingFolderName'
    }

    $inputDir = Resolve-AnyDrivePath -RequestedPath (Resolve-BoundOrDefault -ParameterName 'SourceDir' -EnvironmentNames @('SAILPOINT_SOURCE_DIR') -ConfigValue $sourceConfigured -DefaultValue 'Source Input') -BasePath $utilityPath -SettingName 'SourceDir'
    $processingDir = Resolve-AnyDrivePath -RequestedPath (Resolve-BoundOrDefault -ParameterName 'TemporaryDir' -EnvironmentNames @('SAILPOINT_TEMPORARY_DIR') -ConfigValue $temporaryConfigured -DefaultValue 'Temperary Files') -BasePath $utilityPath -SettingName 'TemporaryDir'

    Get-SailPointCredential -SecretsDir (Join-Path $utilityPath 'Secrets')
    $logRoot = Join-Path $utilityPath $logFolder
    $stateDir = Join-Path $utilityPath $stateFolder
    $jarPath = Join-Path $utilityPath $jar

    foreach ($directory in @($utilityPath, $inputDir, $processingDir, $logRoot, $stateDir)) {
        if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
            try {
                New-Item -Path $directory -ItemType Directory -Force -ErrorAction Stop | Out-Null
            }
            catch {
                Stop-UploadRun -Code 1 -Message "Unable to create directory: $directory. $($_.Exception.Message)"
            }
        }
    }

    Enter-ExecutionLock -Path (Join-Path $utilityPath 'sailpoint_upload.lock')

    Write-Host ''
    Write-Host '============================================================' -ForegroundColor DarkGray
    Write-Host "SailPoint ISC Multi-Source Upload Utility $script:ScriptVersion" -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor DarkGray
    Write-Host ''

    $runStarted = $script:RunStarted
    $runId = $script:RunId
    $runLogDir = $script:RunLogDir

    $matchMode = 'Case-insensitive exact source-name match'
    if ($caseSensitive) {
        $matchMode = 'Case-sensitive exact source-name match'
    }

    Write-Utf8Lines -Path $script:RunSummaryLog -Lines @(
        'CONFIGURATION'
        '----------------------------------------------------------------------'
        "Tenant URL        : $tenant"
        "Input folder      : $inputDir"
        "Processing folder : $processingDir"
        "State folder      : $stateDir"
        "JAR path          : $jarPath"
        "Source match      : $matchMode"
        "Aggregation       : $aggregationMode"
        'Account file      : {SourceName}.csv'
        'Group file        : {SourceName}-group.csv'
        "Credential source : $($script:CredentialSource)"
        'Original retained : YES'
        'Archive used      : NO'
        '----------------------------------------------------------------------'
        ''
    ) -Append

    Write-RunLog -Level INFO -Message "Upload execution started. RunId=$runId CredentialSource=$($script:CredentialSource) Aggregation=$aggregationMode"

    if (-not (Test-Path -LiteralPath $jarPath -PathType Leaf)) {
        Stop-UploadRun -Code 1 -Message "SailPoint upload JAR was not found: $jarPath"
    }

    $javaCommand = Get-Command -Name 'java.exe' -ErrorAction SilentlyContinue
    if ($null -eq $javaCommand) {
        $javaCommand = Get-Command -Name 'java' -ErrorAction SilentlyContinue
    }
    if ($null -eq $javaCommand) {
        Stop-UploadRun -Code 1 -Message 'Java was not found in the system PATH.'
    }

    $javaExecutable = $javaCommand.Source
    try {
        $javaVersion = Get-JavaMajorVersion -JavaExecutable $javaExecutable
    }
    catch {
        Stop-UploadRun -Code 1 -Message "Java version check failed. $(Format-ErrorDetail $_)"
    }
    if ($null -eq $javaVersion.Major) {
        Write-RunLog -Level WARNING -Message "Java version could not be parsed. File Upload Utility 4.1 requires Java 11 or newer. Output: $($javaVersion.Text)"
    }
    elseif ($javaVersion.Major -lt 11) {
        Stop-UploadRun -Code 1 -Message "Java $($javaVersion.Major) is installed. File Upload Utility 4.1 requires Java 11 or newer."
    }
    else {
        Write-RunLog -Level INFO -Message "Java executable: $javaExecutable. Version: $($javaVersion.Text)"
    }

    $accessToken = $null
    try {
        Write-RunLog -Level INFO -Message 'Authenticating to SailPoint ISC.'
        $accessToken = Get-SailPointAccessToken -BaseUrl $tenant -RetryCount $retryCount
        Write-RunLog -Level SUCCESS -Message 'SailPoint ISC authentication completed.'
        Write-RunLog -Level INFO -Message 'Retrieving SailPoint ISC source registry.'
        $sources = @(Get-AllSailPointSources -BaseUrl $tenant -AccessToken $accessToken -PageSize $pageSize -RetryCount $retryCount)
        Write-RunLog -Level SUCCESS -Message "Source registry retrieved. Valid sources found: $($sources.Count)"
    }
    catch {
        if (Test-UploadStop $_) { throw }
        Stop-UploadRun -Code 1 -Message (Format-ErrorDetail $_)
    }
    finally {
        if ($accessToken) {
            $accessToken = $null
        }
    }

    $sourceIndex = New-SourceIndex -Sources $sources -CaseSensitive $caseSensitive
    $duplicateNames = 0
    foreach ($key in @($sourceIndex.Keys)) {
        if ($sourceIndex[$key].Count -gt 1) {
            $duplicateNames++
        }
    }

    if ($duplicateNames -gt 0) {
        Write-RunLog -Level WARNING -Message "Source registry contains $duplicateNames name(s) that match more than one source. Files using those names will fail closed."
    }

    if ($discovery) {
        foreach ($source in $sources) {
            Write-RunLog -Level INFO -Message "SOURCE DISCOVERED | Name=$($source.name) | Id=$($source.id)"
        }
    }

    try {
        $csvFiles = @(
            Get-ChildItem -LiteralPath $inputDir -File -Filter '*.csv' -ErrorAction Stop |
                Sort-Object Name
        )
    }
    catch {
        Stop-UploadRun -Code 1 -Message "Unable to read input CSV files. $($_.Exception.Message)"
    }

    $selectedFiles = New-Object System.Collections.Generic.List[object]
    $ignoredNames = New-Object System.Collections.Generic.List[string]
    foreach ($csvFile in $csvFiles) {
        $plan = Resolve-CsvAggregation -CsvFile $csvFile
        if ($aggregationMode -ne 'both' -and $plan.AggregationType -ne $aggregationMode) {
            $ignoredNames.Add($csvFile.Name)
            continue
        }

        $selectedFiles.Add($csvFile)
    }

    $orderedFiles = @(
        foreach ($csvFile in $selectedFiles) {
            $plan = Resolve-CsvAggregation -CsvFile $csvFile
            $rank = 1
            if ($plan.AggregationType -eq 'group') {
                $rank = 0
            }
            [PSCustomObject]@{
                File = $csvFile
                Rank = $rank
                Name = $csvFile.Name
            }
        }
    ) | Sort-Object Rank, Name
    $selectedFiles = New-Object System.Collections.Generic.List[object]
    foreach ($ordered in $orderedFiles) {
        $selectedFiles.Add($ordered.File)
    }

    $targetComparer = [StringComparer]::OrdinalIgnoreCase
    if ($caseSensitive) {
        $targetComparer = [StringComparer]::Ordinal
    }

    $targetOwners = New-Object System.Collections.Hashtable $targetComparer
    $duplicateTargets = New-Object System.Collections.Hashtable $targetComparer
    foreach ($csvFile in $selectedFiles) {
        $plan = Resolve-CsvAggregation -CsvFile $csvFile
        $targetKey = "$($plan.AggregationType)|$($plan.SourceName)"
        if ($targetOwners.ContainsKey($targetKey)) {
            if ($duplicateTargets.ContainsKey($targetKey)) {
                $duplicateTargets[$targetKey] = "$($duplicateTargets[$targetKey]); $($csvFile.Name)"
            }
            else {
                $duplicateTargets[$targetKey] = "$($targetOwners[$targetKey]); $($csvFile.Name)"
            }

            Write-RunLog -Level ERROR -Message "Duplicate target | ObjectType=$($plan.AggregationType) | Source=$($plan.SourceName) | File=$($csvFile.Name) | Files=$($duplicateTargets[$targetKey])"
        }
        else {
            $targetOwners[$targetKey] = $csvFile.Name
        }
    }

    Write-RunLog -Level INFO -Message "Input CSV files found: $($csvFiles.Count). Selected for '$aggregationMode': $($selectedFiles.Count). Ignored: $($ignoredNames.Count)"
    if ($ignoredNames.Count -gt 0) {
        $shownIgnored = @($ignoredNames)
        if ($shownIgnored.Count -gt 30) {
            $shownIgnored = @($shownIgnored[0..29]) + @('...')
        }
        Write-RunLog -Level INFO -Message "Ignored files do not match aggregation mode '$aggregationMode': $($shownIgnored -join ', ')"
    }
    if ($aggregationMode -eq 'both' -and $selectedFiles.Count -gt 0) {
        Write-RunLog -Level INFO -Message 'Processing order: group files first, then account files.'
    }
    if ($selectedFiles.Count -gt 0) {
        $shownSelected = @($selectedFiles | ForEach-Object { $_.Name })
        if ($shownSelected.Count -gt 40) {
            $shownSelected = @($shownSelected[0..39]) + @('...')
        }
        Write-RunLog -Level INFO -Message "Selected file order: $($shownSelected -join ', ')"
    }

    $settings = [PSCustomObject]@{
        TenantUrl                 = $tenant
        JarPath                   = $jarPath
        JavaExecutable            = $javaExecutable
        ProcessingDir             = $processingDir
        StateDir                  = $stateDir
        RunLogDir                 = $runLogDir
        FileStabilityWaitSeconds  = $stabilitySeconds
        JavaTimeoutMinutes        = $javaTimeout
        ApiRetryCount             = $retryCount
        WhatIf                    = $whatIfEnabled
        MatchMode                 = $matchMode
        PauseAtEnd                = [bool]$pause
        AggregationMode           = $aggregationMode
        DuplicateTargets          = $duplicateTargets
    }

    if ($selectedFiles.Count -eq 0) {
        Write-RunLog -Level WARNING -Message "No CSV files were selected for aggregation mode '$aggregationMode'."
        Write-FinalSummary -Settings $settings -Started $runStarted -FileCount 0 -ExitCode 0
        return 0
    }

    $position = 0
    foreach ($csvFile in $selectedFiles) {
        $position++
        try {
            $oneResult = Invoke-OneCsvFile -CsvFile $csvFile -Position $position -Total $selectedFiles.Count -SourceIndex $sourceIndex -Settings $settings
        }
        catch {
            $safeError = Format-ErrorDetail $_
            Write-RunLog -Level ERROR -Message "FAILED | File=$($csvFile.Name) | Reason=$safeError | The run will continue with the next file."
            $script:FailureCount++
            continue
        }

        switch ($oneResult.Status) {
            'SUCCESS' { $script:SuccessCount++ }
            'FAILED' { $script:FailureCount++ }
            'SKIPPED' { $script:SkippedCount++ }
            'UNCHANGED' { $script:UnchangedCount++ }
            'WHATIF' { $script:WhatIfCount++ }
            default { $script:FailureCount++ }
        }

        if ($oneResult.Status -eq 'SKIPPED') {
            Write-RunLog -Level WARNING -Message "SKIPPED | ObjectType=$($oneResult.AggregationType) | Source=$($oneResult.SourceName) | File=$($csvFile.Name) | Reason=$($oneResult.Reason) | Log=$([IO.Path]::GetFileName($oneResult.AggregationLog))"
        }
        elseif ($oneResult.Status -eq 'UNCHANGED') {
            Write-RunLog -Level INFO -Message "UNCHANGED | ObjectType=$($oneResult.AggregationType) | Source=$($oneResult.SourceName) | File=$($csvFile.Name) | Reason=$($oneResult.Reason) | Log=$([IO.Path]::GetFileName($oneResult.AggregationLog))"
        }
        elseif ($oneResult.Status -eq 'WHATIF') {
            Write-RunLog -Level INFO -Message "WHATIF | ObjectType=$($oneResult.AggregationType) | Source=$($oneResult.SourceName) | SourceId=$($oneResult.SourceId) | File=$($csvFile.Name) | Reason=$($oneResult.Reason)"
        }
        elseif ($oneResult.Status -eq 'FAILED' -and -not $oneResult.SummaryLogged) {
            Write-RunLog -Level ERROR -Message "FAILED | ObjectType=$($oneResult.AggregationType) | Source=$($oneResult.SourceName) | SourceId=$($oneResult.SourceId) | File=$($csvFile.Name) | Reason=$($oneResult.Reason) | Log=$([IO.Path]::GetFileName($oneResult.AggregationLog))"
        }

        if ($oneResult.Status -eq 'FAILED') {
            Write-RunLog -Level INFO -Message "Continuing with the next file after a failure. File=$($csvFile.Name) Position=$position of $($selectedFiles.Count)"
        }
    }

    $accounted = $script:SuccessCount + $script:FailureCount + $script:SkippedCount + $script:UnchangedCount + $script:WhatIfCount
    if ($accounted -ne $selectedFiles.Count) {
        $script:FailureCount++
        Write-RunLog -Level ERROR -Message "Outcome totals ($accounted) do not match the CSV count ($($selectedFiles.Count))."
    }

    Remove-ExpiredRunLogs -LogRoot $logRoot -Keep $runsToKeepValue

    $exitCode = 0
    if ($script:FailureCount -gt 0) {
        $exitCode = 1
    }
    elseif ($script:CleanupWarningCount -gt 0) {
        $exitCode = 4
    }

    try {
        Write-FinalSummary -Settings $settings -Started $runStarted -FileCount $selectedFiles.Count -ExitCode $exitCode
    }
    catch {
        Write-Host "Final summary could not be written. $(Format-ErrorDetail $_)" -ForegroundColor Yellow
        if ($exitCode -eq 0) {
            $exitCode = 4
        }
    }

    if ($exitCode -eq 1) {
        Write-Host ''
        Write-Host 'Execution completed with one or more failures.' -ForegroundColor Red
        Write-Host "Review the summary, error, and aggregation logs in $($script:RunLogDir)" -ForegroundColor Yellow
    }
    elseif ($exitCode -eq 4) {
        Write-Host ''
        Write-Host 'Uploads completed, but cleanup produced warnings.' -ForegroundColor Yellow
    }
    else {
        Write-Host ''
        Write-Host 'All eligible changed files were processed successfully.' -ForegroundColor Green
    }

    return $exitCode
}

$script:PauseRequested = $false
$script:BoundParameters = @{}
foreach ($boundName in @($PSBoundParameters.Keys)) {
    $script:BoundParameters[$boundName] = $PSBoundParameters[$boundName]
}

try {
    $script:RequestedExitCode = Start-SailPointUpload
}
catch {
    if (-not (Test-UploadStop $_)) {
        $cancelled = $_.Exception -is [System.Management.Automation.PipelineStoppedException]
        $safeError = Format-ErrorDetail $_
        if ($cancelled) {
            $safeError = "Run cancelled by the operator. $safeError"
        }

        Ensure-RunLog
        if ($script:RunSummaryLog) {
            Write-RunLog -Level ERROR -Message "Unhandled error: $safeError"
            if ($_.ScriptStackTrace) {
                Write-RunLog -Level ERROR -Message "STACK | $(ConvertTo-SingleLine (Hide-SecretText $_.ScriptStackTrace))"
            }
        }
        else {
            Write-Host "Unhandled error: $safeError" -ForegroundColor Red
        }

        $script:RequestedExitCode = 1
    }
}
finally {
    if ($script:ActiveJavaProcess -and -not $script:ActiveJavaProcess.HasExited) {
        try { $script:ActiveJavaProcess.Kill() } catch { }
    }

    Exit-ExecutionLock
    Clear-UploadSecrets

    if ($script:PauseRequested -and [Environment]::UserInteractive) {
        $redirected = $false
        try { $redirected = [Console]::IsInputRedirected } catch { $redirected = $false }
        if (-not $redirected) {
            Write-Host ''
            try {
                Read-Host -Prompt 'Execution finished. Press ENTER to close this window' | Out-Null
            }
            catch {
                Write-Host "Pause prompt could not be read. $(ConvertTo-SingleLine $_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }
}

exit $script:RequestedExitCode
