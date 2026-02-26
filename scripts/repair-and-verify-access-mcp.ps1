[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "Medium")]
param(
    [string]$DatabasePath = $(if ($env:ACCESS_DATABASE_PATH) { $env:ACCESS_DATABASE_PATH } else { "$env:USERPROFILE\Documents\MyDatabase.accdb" }),
    [switch]$UpdateConfigs,
    [switch]$UpdateCodexConfig,
    [switch]$UpdateClaudeConfig,
    [switch]$SkipCleanup,
    [switch]$SkipTrustedLocation,
    [switch]$SkipRegression
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

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
    $regex = [regex]::new('("mcpServers"\s*:\s*\{[\s\S]*?"access-mcp"\s*:\s*\{[\s\S]*?"command"\s*:\s*)"[^"]*"', [System.Text.RegularExpressions.RegexOptions]::None)
    if (-not $regex.IsMatch($content)) {
        return [pscustomobject]@{
            Path   = $ConfigPath
            Status = "not-found"
            Detail = "mcpServers.access-mcp.command not found."
        }
    }

    $escapedPath = $SelectedBinary.Replace('\', '\\')
    $updatedContent = $regex.Replace(
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

    if (-not $script:PSCmdlet.ShouldProcess($ConfigPath, "Update mcpServers.access-mcp.command")) {
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
        Detail = "Updated mcpServers.access-mcp.command."
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
    (Join-Path $repoRoot "mcp-server-official-x86\MS.Access.MCP.Official.exe")
)

$existingCandidates = @($candidatePaths | Where-Object { Test-Path -LiteralPath $_ })
if ($existingCandidates.Count -eq 0) {
    throw "No candidate binaries found. Checked: $($candidatePaths -join ', ')"
}

Write-Host "Probing candidate binaries via JSON-RPC connect_access..."
$probeResults = New-Object 'System.Collections.Generic.List[object]'
foreach ($candidate in $existingCandidates) {
    Write-Host "Probe: $candidate"
    $result = Invoke-ConnectProbe -ExePath $candidate -DbPath $resolvedDatabasePath
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

$shouldUpdateCodex = $UpdateConfigs -or $UpdateCodexConfig
$shouldUpdateClaude = $UpdateConfigs -or $UpdateClaudeConfig
$configResults = New-Object 'System.Collections.Generic.List[object]'

if ($shouldUpdateCodex) {
    $codexPath = Join-Path $env:USERPROFILE ".codex\config.toml"
    $configResults.Add((Update-CodexConfigCommand -ConfigPath $codexPath -SelectedBinary $selectedBinary))
}
if ($shouldUpdateClaude) {
    $claudePath = Join-Path $env:USERPROFILE ".claude.json"
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

Write-Host ("Regression status: {0}" -f $summary.RegressionStatus)
if ($null -ne $summary.RegressionExitCode) {
    Write-Host ("Regression code  : {0}" -f $summary.RegressionExitCode)
}

if ($summary.RegressionStatus -eq "FAIL") {
    throw "Regression failed with exit code $($summary.RegressionExitCode)."
}
