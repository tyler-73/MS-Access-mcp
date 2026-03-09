param(
    [string]$ServerExe = "$PSScriptRoot\..\mcp-server-official-x64\MS.Access.MCP.Official.exe",
    [string]$DatabasePath = "$env:USERPROFILE\Documents\MyDatabase.accdb"
)

$ErrorActionPreference = "Stop"

$dialogWatcherPath = Join-Path $PSScriptRoot "_dialog_watcher.ps1"
if (-not $PSScriptRoot) {
    $dialogWatcherPath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "_dialog_watcher.ps1"
}
$script:DialogWatcherAvailable = $false
if (Test-Path $dialogWatcherPath) { . $dialogWatcherPath; $script:DialogWatcherAvailable = $true }

if (-not (Test-Path $ServerExe -ErrorAction SilentlyContinue)) {
    $fallbackRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
    $fallbackExe  = Join-Path $fallbackRoot "..\mcp-server-official-x64\MS.Access.MCP.Official.exe"
    if (Test-Path $fallbackExe) { $ServerExe = $fallbackExe }
}

function Decode-McpResult {
    param([object]$Response)
    if ($null -eq $Response) { return $null }
    $text = ""
    try { foreach ($c in $Response.result.content) { $text += $c.text } } catch { return $null }
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    try { return ($text | ConvertFrom-Json) } catch { return @{ raw = $text } }
}

function Add-ToolCall {
    param([System.Collections.Generic.List[object]]$Calls, [int]$Id, [string]$Name, [hashtable]$Arguments = @{})
    $Calls.Add([PSCustomObject]@{ Id = $Id; Name = $Name; Arguments = $Arguments })
}

$script:BatchTimeoutSeconds = 120

function Invoke-McpBatch {
    param([string]$ExePath, [System.Collections.Generic.List[object]]$Calls, [string]$ClientName, [string]$ClientVersion = "1.0")
    if ($script:DialogWatcherAvailable) {
        return Invoke-McpBatchWithTimeout -ExePath $ExePath -Calls $Calls `
            -ClientName $ClientName -ClientVersion $ClientVersion `
            -TimeoutSeconds $script:BatchTimeoutSeconds -SectionName $ClientName
    }
    $jsonLines = New-Object 'System.Collections.Generic.List[string]'
    $jsonLines.Add((@{ jsonrpc="2.0"; id=1; method="initialize"; params=@{
        protocolVersion="2024-11-05"; capabilities=@{}; clientInfo=@{ name=$ClientName; version=$ClientVersion }
    }} | ConvertTo-Json -Depth 40 -Compress))
    $jsonLines.Add((@{ jsonrpc="2.0"; method="notifications/initialized"; params=@{} } | ConvertTo-Json -Depth 20 -Compress))
    foreach ($call in $Calls) {
        $jsonLines.Add((@{ jsonrpc="2.0"; id=$call.Id; method="tools/call"
            params=@{ name=$call.Name; arguments=$call.Arguments }
        } | ConvertTo-Json -Depth 50 -Compress))
    }
    $rawLines = @((($jsonLines -join "`n") | & $ExePath))
    $responses = @{}
    foreach ($line in $rawLines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        try { $parsed = $line | ConvertFrom-Json; if ($null -ne $parsed.id) { $responses[[int]$parsed.id] = $parsed } } catch {}
    }
    return $responses
}

Write-Host "=== Recreating MyDatabase.accdb ==="

# Kill Access and remove old database
Get-Process -Name MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 3
Remove-Item "$DatabasePath".Replace('.accdb', '.laccdb') -Force -ErrorAction SilentlyContinue

# Backup old database
$backupPath = "$DatabasePath".Replace('.accdb', '_backup_pre_recreate.accdb')
if (Test-Path $DatabasePath) {
    Copy-Item $DatabasePath $backupPath -Force
    Write-Host "Backed up to: $backupPath"
    Remove-Item $DatabasePath -Force
}

# Create fresh database
$calls = New-Object 'System.Collections.Generic.List[object]'
Add-ToolCall -Calls $calls -Id 1 -Name "create_database" -Arguments @{ database_path = $DatabasePath }
Add-ToolCall -Calls $calls -Id 2 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }

# Create the parent table needed for relationship tests
Add-ToolCall -Calls $calls -Id 3 -Name "create_table" -Arguments @{
    table_name = "mcp_parent_table"
    fields = @(
        @{ name = "id"; type = "LONG"; size = 0; required = $true; allow_zero_length = $false }
    )
}

# Verify VBA works
Add-ToolCall -Calls $calls -Id 4 -Name "create_module" -Arguments @{ module_name = "TestVBA" }
Add-ToolCall -Calls $calls -Id 5 -Name "set_vba_code" -Arguments @{
    project_name = ""; module_name = "TestVBA"
    code = "Option Explicit`nPublic Sub TestPing()`n    Debug.Print ""Pong""`nEnd Sub"
}
Add-ToolCall -Calls $calls -Id 6 -Name "delete_module" -Arguments @{ project_name = ""; module_name = "TestVBA" }
Add-ToolCall -Calls $calls -Id 7 -Name "disconnect_access" -Arguments @{}
Add-ToolCall -Calls $calls -Id 8 -Name "close_access" -Arguments @{}

$responses = Invoke-McpBatch -ExePath $ServerExe -Calls $calls -ClientName "recreate-db"

Write-Host ""
foreach ($id in @(1,2,3,4,5,6,7,8)) {
    $decoded = Decode-McpResult -Response $responses[[int]$id]
    $label = switch($id) { 1{"create_db"} 2{"connect"} 3{"create_parent_table"} 4{"create_module"} 5{"set_vba_code"} 6{"delete_module"} 7{"disconnect"} 8{"close_access"} }
    if ($null -eq $decoded) { Write-Host "$label`: NULL" }
    elseif ($decoded.success -eq $false) { Write-Host "$label`: ERROR - $($decoded.error)" }
    else { Write-Host "$label`: OK" }
}

Write-Host "`nDone. Fresh MyDatabase.accdb created."
