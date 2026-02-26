[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "Medium")]
param(
    [string]$DatabasePath = $(if ($env:ACCESS_DATABASE_PATH) { $env:ACCESS_DATABASE_PATH } else { "$env:USERPROFILE\Documents\MyDatabase.accdb" }),
    [string]$ClaudeConfigPath = "",
    [switch]$UpdateConfigs,
    [switch]$UpdateCodexConfig,
    [switch]$UpdateClaudeConfig,
    [switch]$SkipCleanup,
    [switch]$SkipTrustedLocation,
    [switch]$SkipRegression,
    [switch]$AllowUnvalidatedBinary,
    [switch]$AllowX86Fallback,
    [switch]$AllowManifestHeadMismatch,
    [switch]$RequireRegressionBackedManifest,
    [switch]$AllowNonRegressionManifest
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
$validationManifestName = "release-validation.json"

function Get-RepoRoot {
    return (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}

function Get-FullPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $Path))
}

function Resolve-ClaudeConfigPath {
    param(
        [Parameter(Mandatory = $false)]
        [string]$Path
    )

    if (-not [string]::IsNullOrWhiteSpace($Path)) {
        return Get-FullPath -Path $Path
    }

    $candidatePaths = New-Object 'System.Collections.Generic.List[string]'
    if (-not [string]::IsNullOrWhiteSpace($env:APPDATA)) {
        $candidatePaths.Add((Join-Path $env:APPDATA "Claude\claude_desktop_config.json"))
    }
    if (-not [string]::IsNullOrWhiteSpace($env:USERPROFILE)) {
        $candidatePaths.Add((Join-Path $env:USERPROFILE ".claude.json"))
    }

    foreach ($candidatePath in $candidatePaths) {
        if (Test-Path -LiteralPath $candidatePath) {
            return $candidatePath
        }
    }

    if ($candidatePaths.Count -eq 0) {
        throw "Unable to determine default Claude config path because APPDATA and USERPROFILE are both unavailable."
    }

    return $candidatePaths[0]
}

function Get-TrimmedString {
    param([object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    return $text.Trim()
}

function Get-GitHeadCommit {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    try {
        $commit = (& git -C $RepoRoot rev-parse HEAD 2>$null)
        if ($LASTEXITCODE -ne 0) {
            return $null
        }

        return (Get-TrimmedString -Value $commit)
    }
    catch {
        return $null
    }
}

function Read-ValidationManifest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath
    )

    if (-not (Test-Path -LiteralPath $ManifestPath)) {
        return [pscustomobject]@{
            Success  = $false
            Manifest = $null
            Error    = "Validation manifest not found."
        }
    }

    try {
        $rawManifest = Get-Content -LiteralPath $ManifestPath -Raw -ErrorAction Stop
        $manifest = $rawManifest | ConvertFrom-Json -ErrorAction Stop
        return [pscustomobject]@{
            Success  = $true
            Manifest = $manifest
            Error    = $null
        }
    }
    catch {
        return [pscustomobject]@{
            Success  = $false
            Manifest = $null
            Error    = $_.Exception.Message
        }
    }
}

function Get-ManifestBooleanValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Manifest,
        [Parameter(Mandatory = $true)]
        [string]$PropertyName
    )

    $property = $Manifest.PSObject.Properties[$PropertyName]
    if ($null -eq $property) {
        return [pscustomobject]@{
            HasValue = $false
            Value    = $false
        }
    }

    $value = $property.Value
    if ($value -is [bool]) {
        return [pscustomobject]@{
            HasValue = $true
            Value    = [bool]$value
        }
    }

    if ($value -is [string]) {
        $parsed = $false
        if ([bool]::TryParse($value, [ref]$parsed)) {
            return [pscustomobject]@{
                HasValue = $true
                Value    = $parsed
            }
        }
    }

    return [pscustomobject]@{
        HasValue = $false
        Value    = $false
    }
}

function Test-ValidationManifestSelectionPolicy {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath,
        [string]$ExpectedGitCommit,
        [switch]$EnforceGitCommitMatch,
        [switch]$EnforceRegressionBackedManifest
    )

    $readResult = Read-ValidationManifest -ManifestPath $ManifestPath
    if (-not $readResult.Success) {
        return [pscustomobject]@{
            IsValid            = $false
            Reason             = "Could not read validation manifest: $($readResult.Error)"
            ManifestPath       = $ManifestPath
            ManifestGitCommit  = $null
            RegressionBacked   = $false
            SmokeTestPassed    = $false
            RegressionRun      = $false
            RegressionPassed   = $false
            ValidationManifest = $null
        }
    }

    $manifest = $readResult.Manifest
    $manifestGitCommitProperty = $manifest.PSObject.Properties["git_commit"]
    $manifestGitCommit = if ($null -ne $manifestGitCommitProperty) {
        Get-TrimmedString -Value $manifestGitCommitProperty.Value
    }
    else {
        $null
    }

    $smokeTestPassedProperty = Get-ManifestBooleanValue -Manifest $manifest -PropertyName "smoke_test_passed"
    $regressionRunProperty = Get-ManifestBooleanValue -Manifest $manifest -PropertyName "regression_run"
    $regressionPassedProperty = Get-ManifestBooleanValue -Manifest $manifest -PropertyName "regression_passed"
    $regressionBacked = if ($regressionPassedProperty.HasValue) {
        $regressionPassedProperty.Value
    }
    elseif ($regressionRunProperty.HasValue) {
        $regressionRunProperty.Value
    }
    else {
        $false
    }

    if ($EnforceGitCommitMatch) {
        if ([string]::IsNullOrWhiteSpace($ExpectedGitCommit)) {
            return [pscustomobject]@{
                IsValid            = $false
                Reason             = "Git HEAD could not be determined for commit validation."
                ManifestPath       = $ManifestPath
                ManifestGitCommit  = $manifestGitCommit
                RegressionBacked   = $regressionBacked
                SmokeTestPassed    = $smokeTestPassedProperty.Value
                RegressionRun      = $regressionRunProperty.Value
                RegressionPassed   = $regressionPassedProperty.Value
                ValidationManifest = $manifest
            }
        }

        if ([string]::IsNullOrWhiteSpace($manifestGitCommit)) {
            return [pscustomobject]@{
                IsValid            = $false
                Reason             = "Validation manifest is missing git_commit."
                ManifestPath       = $ManifestPath
                ManifestGitCommit  = $manifestGitCommit
                RegressionBacked   = $regressionBacked
                SmokeTestPassed    = $smokeTestPassedProperty.Value
                RegressionRun      = $regressionRunProperty.Value
                RegressionPassed   = $regressionPassedProperty.Value
                ValidationManifest = $manifest
            }
        }

        if (-not [string]::Equals($manifestGitCommit, $ExpectedGitCommit, [System.StringComparison]::OrdinalIgnoreCase)) {
            return [pscustomobject]@{
                IsValid            = $false
                Reason             = "Validation manifest git_commit '$manifestGitCommit' does not match repo HEAD '$ExpectedGitCommit'."
                ManifestPath       = $ManifestPath
                ManifestGitCommit  = $manifestGitCommit
                RegressionBacked   = $regressionBacked
                SmokeTestPassed    = $smokeTestPassedProperty.Value
                RegressionRun      = $regressionRunProperty.Value
                RegressionPassed   = $regressionPassedProperty.Value
                ValidationManifest = $manifest
            }
        }
    }

    if ($EnforceRegressionBackedManifest -and (-not $regressionBacked)) {
        return [pscustomobject]@{
            IsValid            = $false
            Reason             = "Strict mode requires a regression-backed validation manifest (regression_run/regression_passed=true)."
            ManifestPath       = $ManifestPath
            ManifestGitCommit  = $manifestGitCommit
            RegressionBacked   = $regressionBacked
            SmokeTestPassed    = $smokeTestPassedProperty.Value
            RegressionRun      = $regressionRunProperty.Value
            RegressionPassed   = $regressionPassedProperty.Value
            ValidationManifest = $manifest
        }
    }

    return [pscustomobject]@{
        IsValid            = $true
        Reason             = $null
        ManifestPath       = $ManifestPath
        ManifestGitCommit  = $manifestGitCommit
        RegressionBacked   = $regressionBacked
        SmokeTestPassed    = $smokeTestPassedProperty.Value
        RegressionRun      = $regressionRunProperty.Value
        RegressionPassed   = $regressionPassedProperty.Value
        ValidationManifest = $manifest
    }
}

function Stop-StaleAccessProcesses {
    $stopped = 0
    $names = @("MS.Access.MCP.Official", "MSACCESS")
    $processes = @(Get-Process -Name $names -ErrorAction SilentlyContinue)
    foreach ($process in $processes) {
        $target = "{0} (PID {1})" -f $process.ProcessName, $process.Id
        if ($script:PSCmdlet.ShouldProcess($target, "Stop process")) {
            try {
                Stop-Process -Id $process.Id -Force -ErrorAction Stop
                $stopped++
                Write-Host "Stopped process: $target"
            }
            catch {
                Write-Warning ("Failed to stop process {0}: {1}" -f $process.Id, $_.Exception.Message)
            }
        }
    }

    return $stopped
}

function Remove-LockFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DbPath
    )

    $dbDir = Split-Path -Path $DbPath -Parent
    $dbName = [System.IO.Path]::GetFileNameWithoutExtension($DbPath)
    $lockFile = Join-Path $dbDir ($dbName + ".laccdb")

    if (-not (Test-Path -LiteralPath $lockFile)) {
        return [pscustomobject]@{
            Path    = $lockFile
            Existed = $false
            Removed = $false
        }
    }

    $removed = $false
    if ($script:PSCmdlet.ShouldProcess($lockFile, "Remove stale lock file")) {
        try {
            Remove-Item -LiteralPath $lockFile -Force -ErrorAction Stop
            $removed = $true
            Write-Host "Removed lock file: $lockFile"
        }
        catch {
            Write-Warning ("Failed to remove lock file {0}: {1}" -f $lockFile, $_.Exception.Message)
        }
    }

    return [pscustomobject]@{
        Path    = $lockFile
        Existed = $true
        Removed = $removed
    }
}

function Get-OfficeVersions {
    $versions = New-Object 'System.Collections.Generic.List[string]'
    $officeRoot = "HKCU:\Software\Microsoft\Office"
    if (Test-Path -LiteralPath $officeRoot) {
        foreach ($subKey in Get-ChildItem -LiteralPath $officeRoot -ErrorAction SilentlyContinue) {
            if ($subKey.PSChildName -match '^\d+\.\d+$') {
                $versions.Add($subKey.PSChildName)
            }
        }
    }

    if ($versions.Count -eq 0) {
        $versions.Add("16.0")
    }
    elseif (-not ($versions.Contains("16.0"))) {
        $versions.Add("16.0")
    }

    return @($versions | Sort-Object -Unique)
}

function Ensure-AccessTrustedLocation {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderPath
    )

    $normalizedFolder = $FolderPath.TrimEnd('\') + '\'
    $versions = Get-OfficeVersions
    $results = New-Object 'System.Collections.Generic.List[object]'

    foreach ($version in $versions) {
        $trustedRoot = "HKCU:\Software\Microsoft\Office\$version\Access\Security\Trusted Locations"
        $status = "skipped"
        $locationKey = $null
        $message = $null

        try {
            if (-not $script:PSCmdlet.ShouldProcess($trustedRoot, "Ensure trusted location for $normalizedFolder")) {
                $results.Add([pscustomobject]@{
                        Version     = $version
                        TrustedRoot = $trustedRoot
                        LocationKey = $locationKey
                        Status      = "whatif"
                        Message     = "Skipped by WhatIf/Confirm"
                    })
                continue
            }

            if (-not (Test-Path -LiteralPath $trustedRoot)) {
                New-Item -Path $trustedRoot -Force | Out-Null
            }

            $matchingLocation = $null
            $existingLocationKeys = @(Get-ChildItem -LiteralPath $trustedRoot -ErrorAction SilentlyContinue)
            foreach ($location in $existingLocationKeys) {
                $props = Get-ItemProperty -LiteralPath $location.PSPath -ErrorAction SilentlyContinue
                if ($null -eq $props) {
                    continue
                }

                $existingPath = [string]$props.Path
                if ([string]::IsNullOrWhiteSpace($existingPath)) {
                    continue
                }

                $normalizedExisting = $existingPath.TrimEnd('\') + '\'
                if ([string]::Equals($normalizedExisting, $normalizedFolder, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $matchingLocation = $location
                    break
                }
            }

            if ($null -ne $matchingLocation) {
                $locationKey = "HKCU:\Software\Microsoft\Office\$version\Access\Security\Trusted Locations\$($matchingLocation.PSChildName)"
                $status = "updated"
            }
            else {
                $maxIndex = 0
                foreach ($location in $existingLocationKeys) {
                    if ($location.PSChildName -match '^Location(\d+)$') {
                        $index = [int]$Matches[1]
                        if ($index -gt $maxIndex) {
                            $maxIndex = $index
                        }
                    }
                }

                $newLocationName = "Location{0:D2}" -f ($maxIndex + 1)
                $locationKey = "HKCU:\Software\Microsoft\Office\$version\Access\Security\Trusted Locations\$newLocationName"
                New-Item -Path $locationKey -Force | Out-Null
                $status = "created"
            }

            New-ItemProperty -Path $locationKey -Name "Path" -PropertyType String -Value $normalizedFolder -Force | Out-Null
            New-ItemProperty -Path $locationKey -Name "AllowSubfolders" -PropertyType DWord -Value 1 -Force | Out-Null
            New-ItemProperty -Path $locationKey -Name "Description" -PropertyType String -Value "MS-Access-mcp hardening script" -Force | Out-Null
            New-ItemProperty -Path $locationKey -Name "Date" -PropertyType String -Value (Get-Date -Format "yyyy-MM-dd HH:mm:ss") -Force | Out-Null

            $results.Add([pscustomobject]@{
                    Version     = $version
                    TrustedRoot = $trustedRoot
                    LocationKey = $locationKey
                    Status      = $status
                    Message     = $message
                })
        }
        catch {
            $results.Add([pscustomobject]@{
                    Version     = $version
                    TrustedRoot = $trustedRoot
                    LocationKey = $locationKey
                    Status      = "failed"
                    Message     = $_.Exception.Message
                })
            Write-Warning ("Trusted location update failed for Office {0}: {1}" -f $version, $_.Exception.Message)
        }
    }

    return $results.ToArray()
}

function Decode-McpResult {
    param([object]$Response)

    if ($null -eq $Response) {
        return $null
    }

    $resultObject = $Response.PSObject.Properties["result"]
    if ($null -eq $resultObject) {
        return $null
    }

    $resultValue = $resultObject.Value
    if ($null -eq $resultValue) {
        return $null
    }

    $structuredContentProperty = $resultValue.PSObject.Properties["structuredContent"]
    if ($null -ne $structuredContentProperty -and $null -ne $structuredContentProperty.Value) {
        return $structuredContentProperty.Value
    }

    $contentProperty = $resultValue.PSObject.Properties["content"]
    if ($null -ne $contentProperty -and $null -ne $contentProperty.Value) {
        $contentItems = @($contentProperty.Value)
        if ($contentItems.Count -eq 0) {
            return $null
        }

        $firstItem = $contentItems[0]
        if ($null -eq $firstItem) {
            return $null
        }

        $textProperty = $firstItem.PSObject.Properties["text"]
        if ($null -eq $textProperty) {
            return $null
        }

        $text = [string]$textProperty.Value
        if ([string]::IsNullOrWhiteSpace($text)) {
            return $null
        }

        try {
            return $text | ConvertFrom-Json
        }
        catch {
            return $text
        }
    }

    return $resultValue
}

function Test-ConnectSuccess {
    param(
        [object]$ConnectResponse,
        [object]$Decoded
    )

    if ($null -eq $ConnectResponse) {
        return $false
    }

    $connectErrorProperty = $ConnectResponse.PSObject.Properties["error"]
    if ($null -ne $connectErrorProperty -and $null -ne $connectErrorProperty.Value) {
        return $false
    }

    if ($null -eq $Decoded) {
        return $false
    }

    if ($Decoded -is [string]) {
        return ([string]::IsNullOrWhiteSpace($Decoded) -eq $false) -and ($Decoded -notmatch '(?i)\bfail|\berror\b')
    }

    $props = @($Decoded.PSObject.Properties.Name)
    $decodedErrorProperty = $Decoded.PSObject.Properties["error"]
    if ($null -ne $decodedErrorProperty -and $null -ne $decodedErrorProperty.Value -and [string]::IsNullOrWhiteSpace([string]$decodedErrorProperty.Value) -eq $false) {
        return $false
    }

    $decodedSuccessProperty = $Decoded.PSObject.Properties["success"]
    if ($null -ne $decodedSuccessProperty -and $decodedSuccessProperty.Value -is [bool]) {
        return [bool]$decodedSuccessProperty.Value
    }

    $decodedConnectedProperty = $Decoded.PSObject.Properties["connected"]
    if ($null -ne $decodedConnectedProperty -and $decodedConnectedProperty.Value -is [bool]) {
        return [bool]$decodedConnectedProperty.Value
    }

    return $false
}

function Invoke-ConnectProbe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExePath,
        [Parameter(Mandatory = $true)]
        [string]$DbPath
    )

    $initializeRequest = @{
        jsonrpc = "2.0"
        id = 1
        method = "initialize"
        params = @{
            protocolVersion = "2024-11-05"
            capabilities = @{}
            clientInfo = @{
                name = "repair-and-verify"
                version = "1.0"
            }
        }
    } | ConvertTo-Json -Depth 20 -Compress

    $connectRequest = @{
        jsonrpc = "2.0"
        id = 2
        method = "tools/call"
        params = @{
            name = "connect_access"
            arguments = @{
                database_path = $DbPath
            }
        }
    } | ConvertTo-Json -Depth 20 -Compress

    $disconnectRequest = @{
        jsonrpc = "2.0"
        id = 3
        method = "tools/call"
        params = @{
            name = "disconnect_access"
            arguments = @{}
        }
    } | ConvertTo-Json -Depth 20 -Compress

    $payload = @($initializeRequest, $connectRequest, $disconnectRequest) -join "`n"
    $nativePrefVar = Get-Variable -Name PSNativeCommandUseErrorActionPreference -ErrorAction SilentlyContinue
    if ($nativePrefVar) {
        $previousNativePreference = $nativePrefVar.Value
        $PSNativeCommandUseErrorActionPreference = $false
    }

    try {
        $rawOutput = @($payload | & $ExePath 2>&1)
    }
    finally {
        if ($nativePrefVar) {
            $PSNativeCommandUseErrorActionPreference = $previousNativePreference
        }
    }

    $exitCode = $LASTEXITCODE
    $responses = @{}
    foreach ($line in $rawOutput) {
        if ([string]::IsNullOrWhiteSpace([string]$line)) {
            continue
        }

        try {
            $parsed = ([string]$line) | ConvertFrom-Json
            if ($null -ne $parsed.id) {
                $responses[[int]$parsed.id] = $parsed
            }
        }
        catch {
            # Ignore non-JSON output lines while probing.
        }
    }

    $connectResponse = $null
    if ($responses.ContainsKey(2)) {
        $connectResponse = $responses[2]
    }

    $decoded = Decode-McpResult -Response $connectResponse
    $success = Test-ConnectSuccess -ConnectResponse $connectResponse -Decoded $decoded
    $errorText = $null

    if (-not $success) {
        if ($null -eq $connectResponse) {
            $errorText = "No connect_access JSON-RPC response (id=2)."
        }
        else {
            $connectErrorProperty = $connectResponse.PSObject.Properties["error"]
            if ($null -ne $connectErrorProperty -and $null -ne $connectErrorProperty.Value) {
                $errorText = [string]$connectErrorProperty.Value
            }
            elseif ($decoded -is [string]) {
                $errorText = $decoded
            }
            elseif ($null -ne $decoded) {
                try {
                    $errorText = ($decoded | ConvertTo-Json -Depth 20 -Compress)
                }
                catch {
                    $errorText = [string]$decoded
                }
            }
            else {
                $errorText = "connect_access returned an unreadable payload."
            }
        }
    }

    return [pscustomobject]@{
        ExePath      = $ExePath
        ExitCode     = $exitCode
        Success      = $success
        Error        = $errorText
        RawOutput    = @($rawOutput)
        ConnectReply = $connectResponse
    }
}

function Invoke-InitializeSmokeProbe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExePath,
        [Parameter(Mandatory = $true)]
        [string]$SmokeScriptPath
    )

    if (-not (Test-Path -LiteralPath $SmokeScriptPath)) {
        return [pscustomobject]@{
            ExePath   = $ExePath
            ExitCode  = -1
            Success   = $false
            Error     = "Smoke script not found: $SmokeScriptPath"
            RawOutput = @()
        }
    }

    $powershellExe = Get-PowerShellHostExecutable
    $nativePrefVar = Get-Variable -Name PSNativeCommandUseErrorActionPreference -ErrorAction SilentlyContinue
    if ($nativePrefVar) {
        $previousNativePreference = $nativePrefVar.Value
        $PSNativeCommandUseErrorActionPreference = $false
    }

    try {
        $rawOutput = @(& $powershellExe -NoProfile -ExecutionPolicy Bypass -File $SmokeScriptPath -ServerExe $ExePath 2>&1)
    }
    finally {
        if ($nativePrefVar) {
            $PSNativeCommandUseErrorActionPreference = $previousNativePreference
        }
    }

    $exitCode = $LASTEXITCODE
    $success = ($exitCode -eq 0)
    $errorText = $null
    if (-not $success) {
        $errorText = (@($rawOutput) | Select-Object -Last 5) -join [Environment]::NewLine
        if ([string]::IsNullOrWhiteSpace($errorText)) {
            $errorText = "Initialize smoke failed with exit code $exitCode."
        }
    }

    return [pscustomobject]@{
        ExePath   = $ExePath
        ExitCode  = $exitCode
        Success   = $success
        Error     = $errorText
        RawOutput = @($rawOutput)
    }
}

function Update-CodexConfigCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath,
        [Parameter(Mandatory = $true)]
        [string]$SelectedBinary
    )

    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        return [pscustomobject]@{
            Path   = $ConfigPath
            Status = "missing"
            Detail = "File not found."
        }
    }

    $lines = @(Get-Content -LiteralPath $ConfigPath)
    $sectionIndex = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i].Trim() -eq "[mcp_servers.access-mcp]") {
            $sectionIndex = $i
            break
        }
    }

    if ($sectionIndex -lt 0) {
        return [pscustomobject]@{
            Path   = $ConfigPath
            Status = "not-found"
            Detail = "Section [mcp_servers.access-mcp] not found."
        }
    }

    $sectionEnd = $lines.Count
    for ($i = $sectionIndex + 1; $i -lt $lines.Count; $i++) {
        if ($lines[$i].Trim() -match '^\[.+\]$') {
            $sectionEnd = $i
            break
        }
    }

    $commandLineIndex = -1
    for ($i = $sectionIndex + 1; $i -lt $sectionEnd; $i++) {
        if ($lines[$i] -match '^\s*command\s*=') {
            $commandLineIndex = $i
            break
        }
    }

    $newCommandLine = "command = '{0}'" -f $SelectedBinary
    $changed = $false
    if ($commandLineIndex -ge 0) {
        if ($lines[$commandLineIndex] -ne $newCommandLine) {
            $lines[$commandLineIndex] = $newCommandLine
            $changed = $true
        }
    }
    else {
        $before = @($lines[0..$sectionIndex])
        $after = @()
        if ($sectionIndex + 1 -lt $lines.Count) {
            $after = @($lines[($sectionIndex + 1)..($lines.Count - 1)])
        }
        $lines = @($before + $newCommandLine + $after)
        $changed = $true
    }

    if (-not $changed) {
        return [pscustomobject]@{
            Path   = $ConfigPath
            Status = "unchanged"
            Detail = "Command already matches selected binary."
        }
    }

    if (-not $script:PSCmdlet.ShouldProcess($ConfigPath, "Update mcp_servers.access-mcp command")) {
        return [pscustomobject]@{
            Path   = $ConfigPath
            Status = "whatif"
            Detail = "Skipped by WhatIf/Confirm."
        }
    }

    Set-Content -LiteralPath $ConfigPath -Value $lines -Encoding utf8
    return [pscustomobject]@{
        Path   = $ConfigPath
        Status = "updated"
        Detail = "Updated mcp_servers.access-mcp command."
    }
}

function Update-ClaudeConfigCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath,
        [Parameter(Mandatory = $true)]
        [string]$SelectedBinary
    )

    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        return [pscustomobject]@{
            Path   = $ConfigPath
            Status = "missing"
            Detail = "File not found."
        }
    }

    $content = Get-Content -LiteralPath $ConfigPath -Raw
    $patterns = @(
        [pscustomobject]@{
            ServerName = "access-mcp-server"
            Regex      = [regex]::new('("mcpServers"\s*:\s*\{[\s\S]*?"access-mcp-server"\s*:\s*\{[\s\S]*?"command"\s*:\s*)"[^"]*"', [System.Text.RegularExpressions.RegexOptions]::None)
        },
        [pscustomobject]@{
            ServerName = "access-mcp"
            Regex      = [regex]::new('("mcpServers"\s*:\s*\{[\s\S]*?"access-mcp"\s*:\s*\{[\s\S]*?"command"\s*:\s*)"[^"]*"', [System.Text.RegularExpressions.RegexOptions]::None)
        }
    )

    $matchedPattern = $null
    foreach ($pattern in $patterns) {
        if ($pattern.Regex.IsMatch($content)) {
            $matchedPattern = $pattern
            break
        }
    }

    if ($null -eq $matchedPattern) {
        return [pscustomobject]@{
            Path   = $ConfigPath
            Status = "not-found"
            Detail = "mcpServers.access-mcp-server.command or mcpServers.access-mcp.command not found."
        }
    }

    $escapedPath = $SelectedBinary.Replace('\', '\\')
    $updatedContent = $matchedPattern.Regex.Replace(
        $content,
        {
            param($match)
            return $match.Groups[1].Value + '"' + $escapedPath + '"'
        },
        1
    )

    if ($updatedContent -eq $content) {
        return [pscustomobject]@{
            Path   = $ConfigPath
            Status = "unchanged"
            Detail = "Command already matches selected binary."
        }
    }

    if (-not $script:PSCmdlet.ShouldProcess($ConfigPath, ("Update mcpServers.{0}.command" -f $matchedPattern.ServerName))) {
        return [pscustomobject]@{
            Path   = $ConfigPath
            Status = "whatif"
            Detail = "Skipped by WhatIf/Confirm."
        }
    }

    Set-Content -LiteralPath $ConfigPath -Value $updatedContent -Encoding utf8
    return [pscustomobject]@{
        Path   = $ConfigPath
        Status = "updated"
        Detail = ("Updated mcpServers.{0}.command." -f $matchedPattern.ServerName)
    }
}

function Get-PowerShellHostExecutable {
    $pwsh = Get-Command -Name "pwsh" -ErrorAction SilentlyContinue
    if ($pwsh) {
        return $pwsh.Source
    }

    $powershell = Get-Command -Name "powershell" -ErrorAction SilentlyContinue
    if ($powershell) {
        return $powershell.Source
    }

    throw "No PowerShell host executable (pwsh/powershell) was found in PATH."
}

$repoRoot = Get-RepoRoot
$resolvedDatabasePath = Get-FullPath -Path $DatabasePath

if (-not (Test-Path -LiteralPath $resolvedDatabasePath)) {
    throw "Database file not found: $resolvedDatabasePath"
}

Write-Host "Repo root: $repoRoot"
Write-Host "Database: $resolvedDatabasePath"

$summary = [ordered]@{
    ProcessesStopped   = 0
    LockFilePath       = $null
    LockFileExisted    = $false
    LockFileRemoved    = $false
    TrustedCreated     = 0
    TrustedUpdated     = 0
    TrustedFailed      = 0
    TrustedWhatIf      = 0
    SelectedBinary     = $null
    ValidationManifest = $null
    ExpectedGitCommit  = $null
    ManifestGitCommit  = $null
    RegressionBackedManifest = $false
    StrictManifestMode = "disabled"
    RegressionExitCode = $null
    RegressionStatus   = "not-run"
}

if ($SkipCleanup) {
    Write-Warning "Cleanup skip requested, but cleanup is enforced for reliability."
}

Write-Host "Stopping stale Access/MCP processes..."
$summary.ProcessesStopped = Stop-StaleAccessProcesses
$lockResult = Remove-LockFile -DbPath $resolvedDatabasePath
$summary.LockFilePath = $lockResult.Path
$summary.LockFileExisted = $lockResult.Existed
$summary.LockFileRemoved = $lockResult.Removed

if (-not $SkipTrustedLocation) {
    Write-Host "Ensuring Access Trusted Location for database folder..."
    $dbFolder = Split-Path -Path $resolvedDatabasePath -Parent
    $trustedResults = Ensure-AccessTrustedLocation -FolderPath $dbFolder
    $summary.TrustedCreated = @($trustedResults | Where-Object { $_.Status -eq "created" }).Count
    $summary.TrustedUpdated = @($trustedResults | Where-Object { $_.Status -eq "updated" }).Count
    $summary.TrustedFailed = @($trustedResults | Where-Object { $_.Status -eq "failed" }).Count
    $summary.TrustedWhatIf = @($trustedResults | Where-Object { $_.Status -eq "whatif" }).Count
}
else {
    Write-Host "Trusted location step skipped by -SkipTrustedLocation."
}

$candidatePaths = @(
    (Join-Path $repoRoot "mcp-server-official-x64\MS.Access.MCP.Official.exe")
)
if ($AllowX86Fallback) {
    Write-Warning "AllowX86Fallback enabled; including mcp-server-official-x86 candidate."
    $candidatePaths += (Join-Path $repoRoot "mcp-server-official-x86\MS.Access.MCP.Official.exe")
}

$existingCandidates = @($candidatePaths | Where-Object { Test-Path -LiteralPath $_ })
if ($existingCandidates.Count -eq 0) {
    throw "No candidate binaries found. Checked: $($candidatePaths -join ', ')"
}

$enforceGitCommitMatch = -not $AllowManifestHeadMismatch
$headGitCommit = Get-GitHeadCommit -RepoRoot $repoRoot
if ($enforceGitCommitMatch) {
    if ([string]::IsNullOrWhiteSpace($headGitCommit)) {
        throw "Unable to determine repo HEAD commit for validation-manifest checks. Use -AllowManifestHeadMismatch to bypass commit matching."
    }

    Write-Host ("Validation policy: manifest git_commit must match repo HEAD ({0})." -f $headGitCommit)
}
else {
    Write-Warning "AllowManifestHeadMismatch enabled; manifest git_commit mismatch checks are bypassed."
}

$enforceRegressionBackedManifest = $RequireRegressionBackedManifest -and (-not $AllowNonRegressionManifest)
if ($enforceRegressionBackedManifest) {
    Write-Host "Validation policy: strict mode enabled, requiring regression-backed validation manifests."
}
elseif ($RequireRegressionBackedManifest -and $AllowNonRegressionManifest) {
    Write-Warning "AllowNonRegressionManifest override enabled; strict regression-backed manifest requirement is bypassed."
}
elseif ($AllowNonRegressionManifest) {
    Write-Warning "AllowNonRegressionManifest has no effect without -RequireRegressionBackedManifest."
}

$summary.ExpectedGitCommit = $headGitCommit
if ($enforceRegressionBackedManifest) {
    $summary.StrictManifestMode = "enabled"
}
elseif ($RequireRegressionBackedManifest -and $AllowNonRegressionManifest) {
    $summary.StrictManifestMode = "override"
}

$candidateRecords = New-Object 'System.Collections.Generic.List[object]'
foreach ($candidate in $existingCandidates) {
    $candidateDirectory = Split-Path -Path $candidate -Parent
    $manifestPath = Join-Path $candidateDirectory $validationManifestName
    $hasManifest = Test-Path -LiteralPath $manifestPath

    if ($hasManifest) {
        $manifestPolicyResult = Test-ValidationManifestSelectionPolicy `
            -ManifestPath $manifestPath `
            -ExpectedGitCommit $headGitCommit `
            -EnforceGitCommitMatch:$enforceGitCommitMatch `
            -EnforceRegressionBackedManifest:$enforceRegressionBackedManifest

        if (-not $manifestPolicyResult.IsValid) {
            Write-Warning ("Skipping candidate due to validation manifest policy failure: {0} | {1}" -f $candidate, $manifestPolicyResult.Reason)
            continue
        }

        $candidateRecords.Add([pscustomobject]@{
                ExePath = $candidate
                ValidationManifestPath = $manifestPath
                ManifestGitCommit = $manifestPolicyResult.ManifestGitCommit
                ManifestRegressionBacked = $manifestPolicyResult.RegressionBacked
                ValidationManifest = $manifestPolicyResult.ValidationManifest
            })
        continue
    }

    if ($AllowUnvalidatedBinary) {
        Write-Warning "Including unvalidated candidate by request: $candidate"
        $candidateRecords.Add([pscustomobject]@{
                ExePath = $candidate
                ValidationManifestPath = $null
                ManifestGitCommit = $null
                ManifestRegressionBacked = $false
                ValidationManifest = $null
            })
    }
    else {
        Write-Warning "Skipping candidate without validation manifest ($validationManifestName): $candidate"
    }
}

if ($candidateRecords.Count -eq 0) {
    throw "No candidate binaries satisfied validation policy. Default policy requires release-validation.json with git_commit matching HEAD. Overrides: -AllowManifestHeadMismatch (commit check), -AllowUnvalidatedBinary (missing manifest), -AllowNonRegressionManifest (strict-mode regression bypass)."
}

$smokeScriptPath = Join-Path $repoRoot "tests\ci_initialize_smoke.ps1"
Write-Host "Validating candidate binaries (initialize smoke + connect_access)..."
$probeResults = New-Object 'System.Collections.Generic.List[object]'
foreach ($candidateRecord in $candidateRecords) {
    $candidate = $candidateRecord.ExePath
    Write-Host "Initialize smoke: $candidate"
    $smokeResult = Invoke-InitializeSmokeProbe -ExePath $candidate -SmokeScriptPath $smokeScriptPath
    if (-not $smokeResult.Success) {
        Write-Warning ("Initialize smoke failed: {0} | ExitCode={1} | Reason={2}" -f $candidate, $smokeResult.ExitCode, $smokeResult.Error)
        $probeResults.Add([pscustomobject]@{
                ExePath      = $candidate
                ExitCode     = $smokeResult.ExitCode
                Success      = $false
                Error        = "initialize smoke failed: $($smokeResult.Error)"
                RawOutput    = @($smokeResult.RawOutput)
                ConnectReply = $null
                ValidationManifestPath = $candidateRecord.ValidationManifestPath
                ManifestGitCommit = $candidateRecord.ManifestGitCommit
                ManifestRegressionBacked = $candidateRecord.ManifestRegressionBacked
            })
        continue
    }

    Write-Host "Connect probe: $candidate"
    $result = Invoke-ConnectProbe -ExePath $candidate -DbPath $resolvedDatabasePath
    $result | Add-Member -NotePropertyName ValidationManifestPath -NotePropertyValue $candidateRecord.ValidationManifestPath
    $result | Add-Member -NotePropertyName ManifestGitCommit -NotePropertyValue $candidateRecord.ManifestGitCommit
    $result | Add-Member -NotePropertyName ManifestRegressionBacked -NotePropertyValue $candidateRecord.ManifestRegressionBacked
    $probeResults.Add($result)
    if ($result.Success) {
        Write-Host "Probe success: $candidate"
    }
    else {
        Write-Warning ("Probe failed: {0} | ExitCode={1} | Reason={2}" -f $candidate, $result.ExitCode, $result.Error)
    }
}

$selectedProbe = @($probeResults | Where-Object { $_.Success } | Select-Object -First 1)
if ($selectedProbe.Count -eq 0) {
    Write-Host "No working binary identified. Probe details:"
    foreach ($probe in $probeResults) {
        Write-Host ("- {0} => Success={1}, ExitCode={2}, Error={3}" -f $probe.ExePath, $probe.Success, $probe.ExitCode, $probe.Error)
    }
    throw "Unable to find a working Access MCP binary via connect_access probe."
}

$selectedBinary = $selectedProbe[0].ExePath
$summary.SelectedBinary = $selectedBinary
$summary.ValidationManifest = $selectedProbe[0].ValidationManifestPath
$summary.ManifestGitCommit = $selectedProbe[0].ManifestGitCommit
$summary.RegressionBackedManifest = [bool]$selectedProbe[0].ManifestRegressionBacked

$shouldUpdateCodex = $UpdateConfigs -or $UpdateCodexConfig
$shouldUpdateClaude = $UpdateConfigs -or $UpdateClaudeConfig
$configResults = New-Object 'System.Collections.Generic.List[object]'

if ($shouldUpdateCodex) {
    $codexPath = Join-Path $env:USERPROFILE ".codex\config.toml"
    $configResults.Add((Update-CodexConfigCommand -ConfigPath $codexPath -SelectedBinary $selectedBinary))
}
if ($shouldUpdateClaude) {
    $claudePath = Resolve-ClaudeConfigPath -Path $ClaudeConfigPath
    $configResults.Add((Update-ClaudeConfigCommand -ConfigPath $claudePath -SelectedBinary $selectedBinary))
}

if (-not $SkipRegression) {
    $regressionScript = Join-Path $repoRoot "tests\full_toolset_regression.ps1"
    if (-not (Test-Path -LiteralPath $regressionScript)) {
        throw "Regression script not found: $regressionScript"
    }

    if ($PSCmdlet.ShouldProcess($regressionScript, "Run full toolset regression with selected binary")) {
        Write-Host "Running regression script: $regressionScript"
        $powershellExe = Get-PowerShellHostExecutable
        & $powershellExe -NoProfile -ExecutionPolicy Bypass -File $regressionScript -ServerExe $selectedBinary -DatabasePath $resolvedDatabasePath
        $summary.RegressionExitCode = $LASTEXITCODE
        if ($LASTEXITCODE -eq 0) {
            $summary.RegressionStatus = "PASS"
        }
        else {
            $summary.RegressionStatus = "FAIL"
        }
    }
    else {
        $summary.RegressionStatus = "SKIPPED-WHATIF"
    }
}
else {
    $summary.RegressionStatus = "SKIPPED-BY-PARAMETER"
}

Write-Host ""
Write-Host "=== repair-and-verify-access-mcp summary ==="
Write-Host ("Selected binary : {0}" -f $summary.SelectedBinary)
if (-not [string]::IsNullOrWhiteSpace([string]$summary.ValidationManifest)) {
    Write-Host ("Validation manifest: {0}" -f $summary.ValidationManifest)
    Write-Host ("Manifest git_commit: {0}" -f $(if ([string]::IsNullOrWhiteSpace([string]$summary.ManifestGitCommit)) { "<missing>" } else { $summary.ManifestGitCommit }))
    Write-Host ("Manifest regression-backed: {0}" -f $summary.RegressionBackedManifest)
}
else {
    Write-Host "Validation manifest: none (unvalidated mode)"
}
if (-not $AllowManifestHeadMismatch) {
    Write-Host ("Expected git_commit: {0}" -f $summary.ExpectedGitCommit)
}
else {
    Write-Host "Expected git_commit: bypassed by -AllowManifestHeadMismatch"
}
Write-Host ("Strict manifest mode: {0}" -f $summary.StrictManifestMode)
Write-Host ("Processes stopped: {0}" -f $summary.ProcessesStopped)
Write-Host ("Lock file path   : {0}" -f $summary.LockFilePath)
Write-Host ("Lock file existed: {0}" -f $summary.LockFileExisted)
Write-Host ("Lock file removed: {0}" -f $summary.LockFileRemoved)
Write-Host ("Trusted locations: created={0}, updated={1}, failed={2}, whatif={3}" -f $summary.TrustedCreated, $summary.TrustedUpdated, $summary.TrustedFailed, $summary.TrustedWhatIf)

if ($configResults.Count -gt 0) {
    Write-Host "Config updates:"
    foreach ($configResult in $configResults) {
        Write-Host ("- {0}: {1} ({2})" -f $configResult.Path, $configResult.Status, $configResult.Detail)
    }
}
else {
    Write-Host "Config updates   : skipped"
}

$failedConfigResults = @($configResults | Where-Object { $_.Status -in @("missing", "not-found", "failed") })
if ($failedConfigResults.Count -gt 0) {
    $details = @($failedConfigResults | ForEach-Object { "{0} [{1}]: {2}" -f $_.Path, $_.Status, $_.Detail })
    throw ("One or more config updates failed: {0}" -f ($details -join "; "))
}

Write-Host ("Regression status: {0}" -f $summary.RegressionStatus)
if ($null -ne $summary.RegressionExitCode) {
    Write-Host ("Regression code  : {0}" -f $summary.RegressionExitCode)
}

if ($summary.RegressionStatus -eq "FAIL") {
    throw "Regression failed with exit code $($summary.RegressionExitCode)."
}
