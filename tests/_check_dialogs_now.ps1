. "$PSScriptRoot\_dialog_watcher.ps1"

$procs = @(Get-Process -Name MSACCESS -ErrorAction SilentlyContinue)
if ($procs.Count -eq 0) {
    Write-Host "No MSACCESS.exe processes found."
    exit 0
}

foreach ($proc in $procs) {
    $procId = [uint32]$proc.Id
    Write-Host ("Checking MSACCESS pid={0}" -f $procId)
    $dialogs = [DialogDetector]::FindDialogsForProcess($procId)
    if ($dialogs.Count -eq 0) {
        Write-Host "  No dialogs detected."
    }
    foreach ($d in $dialogs) {
        Write-Host ("  DIALOG: title='{0}' class={1}" -f $d.Title, $d.ClassName)
        foreach ($ct in $d.ChildTexts) {
            Write-Host ("    child: '{0}'" -f $ct)
        }
    }
}
