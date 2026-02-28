Get-Process -Name MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Get-Process -Name 'MS.Access.MCP.Official' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
$lockPath = Join-Path ([System.IO.Path]::GetTempPath()) 'ms-access-mcp-regression.lock'
if (Test-Path $lockPath) {
    Remove-Item $lockPath -Force -ErrorAction SilentlyContinue
    Write-Host "Lock file removed: $lockPath"
}
else {
    Write-Host "No lock file found."
}
Write-Host "Done."
