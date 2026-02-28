# Kill all stale powershell processes except this one, remove lock
$myPid = $PID
Get-Process -Name powershell -ErrorAction SilentlyContinue |
    Where-Object { $_.Id -ne $myPid } |
    ForEach-Object {
        Write-Host ("Killing powershell pid={0} started={1}" -f $_.Id, $_.StartTime)
        Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
    }

Get-Process -Name MSACCESS -ErrorAction SilentlyContinue |
    ForEach-Object {
        Write-Host ("Killing MSACCESS pid={0}" -f $_.Id)
        Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
    }

Get-Process -Name 'MS.Access.MCP.Official' -ErrorAction SilentlyContinue |
    ForEach-Object {
        Write-Host ("Killing MCP server pid={0}" -f $_.Id)
        Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
    }

Start-Sleep -Seconds 2

$lockPath = Join-Path ([System.IO.Path]::GetTempPath()) 'ms-access-mcp-regression.lock'
if (Test-Path $lockPath) {
    Remove-Item $lockPath -Force -ErrorAction SilentlyContinue
    Write-Host "Lock file removed"
} else {
    Write-Host "No lock file found"
}

Start-Sleep -Seconds 1
if (Test-Path $lockPath) {
    Write-Host "WARNING: Lock file still exists after cleanup!"
} else {
    Write-Host "Lock file confirmed gone - safe to run regression"
}
