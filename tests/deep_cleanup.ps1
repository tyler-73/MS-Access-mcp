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

Write-Host "=== Deep VBA Cleanup ==="

# Step 1: List modules
$calls = New-Object 'System.Collections.Generic.List[object]'
Add-ToolCall -Calls $calls -Id 1 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
Add-ToolCall -Calls $calls -Id 2 -Name "get_modules" -Arguments @{}
$r = Invoke-McpBatch -ExePath $ServerExe -Calls $calls -ClientName "deep-cleanup-list"

$d = Decode-McpResult -Response $r[2]
Write-Host "Found modules:"
$orphans = @()
if ($d -and $d.modules) {
    foreach ($m in $d.modules) {
        $name = $m.name
        Write-Host "  $name"
        if ($name -match "^MCP_") { $orphans += $name }
    }
}

if ($orphans.Count -eq 0) {
    Write-Host "No orphaned MCP_ modules found."
} else {
    Write-Host "`nDeleting $($orphans.Count) orphaned modules..."
    $delCalls = New-Object 'System.Collections.Generic.List[object]'
    Add-ToolCall -Calls $delCalls -Id 1 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
    $id = 2
    foreach ($name in $orphans) {
        Add-ToolCall -Calls $delCalls -Id $id -Name "delete_module" -Arguments @{ project_name = ""; module_name = $name }
        $id++
    }
    # Compile and save all modules to persist the deletion
    Add-ToolCall -Calls $delCalls -Id $id -Name "compile_vba" -Arguments @{}
    $id++
    Add-ToolCall -Calls $delCalls -Id $id -Name "disconnect_access" -Arguments @{}
    $id++
    Add-ToolCall -Calls $delCalls -Id $id -Name "close_access" -Arguments @{}

    $dr = Invoke-McpBatch -ExePath $ServerExe -Calls $delCalls -ClientName "deep-cleanup-delete"

    foreach ($did in 2..($orphans.Count + 1)) {
        $dd = Decode-McpResult -Response $dr[$did]
        $modName = $orphans[$did - 2]
        if ($dd.success -eq $false) { Write-Host "  Delete $modName`: ERROR - $($dd.error)" }
        else { Write-Host "  Delete $modName`: OK" }
    }

    $compileId = $orphans.Count + 2
    $dd = Decode-McpResult -Response $dr[$compileId]
    if ($dd.success -eq $false) { Write-Host "  compile_vba: ERROR - $($dd.error)" }
    else { Write-Host "  compile_vba: OK" }
}

# Wait for clean exit
Write-Host "`nWaiting for Access to exit cleanly..."
Start-Sleep -Seconds 10
Get-Process -Name MSACCESS -ErrorAction SilentlyContinue | ForEach-Object {
    Write-Host "  Killing remaining PID $($_.Id)..."
    $_ | Stop-Process -Force -ErrorAction SilentlyContinue
}
Start-Sleep -Seconds 3
Remove-Item "$DatabasePath".Replace('.accdb', '.laccdb') -Force -ErrorAction SilentlyContinue

# Compact to clean up binary data
Write-Host "`nCompacting database..."
$compactDst = "$DatabasePath".Replace('.accdb', '_compacted.accdb')
Remove-Item $compactDst -Force -ErrorAction SilentlyContinue
$engine = New-Object -ComObject 'DAO.DBEngine.120'
$engine.CompactDatabase($DatabasePath, $compactDst)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($engine) | Out-Null
Remove-Item $DatabasePath -Force
Rename-Item $compactDst (Split-Path $DatabasePath -Leaf)
Write-Host "Compact done."

# Test VBA creation
Write-Host "`nTesting VBA creation after cleanup..."
$testCalls = New-Object 'System.Collections.Generic.List[object]'
Add-ToolCall -Calls $testCalls -Id 1 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
Add-ToolCall -Calls $testCalls -Id 2 -Name "get_modules" -Arguments @{}
Add-ToolCall -Calls $testCalls -Id 3 -Name "create_module" -Arguments @{ module_name = "VbaCleanupTest" }
Add-ToolCall -Calls $testCalls -Id 4 -Name "set_vba_code" -Arguments @{
    project_name = ""; module_name = "VbaCleanupTest"
    code = "Option Explicit`nPublic Sub CleanTest()`n    Debug.Print ""Clean""`nEnd Sub"
}
Add-ToolCall -Calls $testCalls -Id 5 -Name "delete_module" -Arguments @{ project_name = ""; module_name = "VbaCleanupTest" }
Add-ToolCall -Calls $testCalls -Id 6 -Name "disconnect_access" -Arguments @{}
$tr = Invoke-McpBatch -ExePath $ServerExe -Calls $testCalls -ClientName "deep-cleanup-test"

$labels = @{ 1="connect"; 2="get_modules"; 3="create_module"; 4="set_vba_code"; 5="delete_module"; 6="disconnect" }
foreach ($tid in ($labels.Keys | Sort-Object)) {
    $td = Decode-McpResult -Response $tr[$tid]
    $lbl = $labels[$tid]
    if ($null -eq $td) { Write-Host "  $lbl`: NULL" }
    elseif ($td.success -eq $false) { Write-Host "  $lbl`: ERROR - $($td.error)" }
    else { Write-Host "  $lbl`: OK" }
}

# Show remaining modules
$modResp = Decode-McpResult -Response $tr[2]
if ($modResp -and $modResp.modules) {
    Write-Host "`nRemaining modules after cleanup:"
    foreach ($m in $modResp.modules) { Write-Host "  $($m.name)" }
}
