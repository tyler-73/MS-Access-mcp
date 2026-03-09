$src = "$env:USERPROFILE\Documents\MyDatabase.accdb"
$dst = "$env:USERPROFILE\Documents\MyDatabase_compact.accdb"
if (Test-Path $dst) { Remove-Item $dst -Force }
$engine = New-Object -ComObject 'DAO.DBEngine.120'
$engine.CompactDatabase($src, $dst)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($engine) | Out-Null
Remove-Item $src -Force
Rename-Item $dst (Split-Path $src -Leaf)
Write-Host "Compact and repair done"
