# _dialog_watcher.ps1 — Modal dialog detection, screenshot capture, and auto-dismiss for access-mcp testing.
# Dot-source this file from test harnesses: . "$PSScriptRoot\_dialog_watcher.ps1"

# Capture this script's own path at load time (before $PSCommandPath changes in dot-source context).
# $MyInvocation.MyCommand.Path is reliable even when dot-sourced.
$script:_DialogWatcherScriptPath = $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($script:_DialogWatcherScriptPath)) {
    $script:_DialogWatcherScriptPath = $MyInvocation.MyCommand.Definition
}

# ── Win32 P/Invoke via inline C# ──────────────────────────────────────────────

Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public class DialogDetector
{
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    public delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool EnumChildWindows(IntPtr hWndParent, EnumChildProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr FindWindowEx(IntPtr hWndParent, IntPtr hWndChildAfter, string lpszClass, string lpszWindow);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    // Constants
    public const uint GW_OWNER = 4;
    public const int GWL_STYLE = -16;
    public const int GWL_EXSTYLE = -20;
    public const int WS_DISABLED = 0x08000000;
    public const int DS_MODALFRAME = 0x80;
    public const int WS_EX_DLGMODALFRAME = 0x00000001;
    public const uint WM_CLOSE = 0x0010;
    public const uint BM_CLICK = 0x00F5;
    public const uint WM_COMMAND = 0x0111;

    public class DialogInfo
    {
        public IntPtr Handle;
        public string Title;
        public string ClassName;
        public uint ProcessId;
        public bool IsDialog;
        public List<string> ChildTexts;
        public List<ChildControlInfo> ChildControls;
    }

    public class ChildControlInfo
    {
        public IntPtr Handle;
        public string Text;
        public string ClassName;
    }

    public static string GetWindowTitle(IntPtr hWnd)
    {
        int length = GetWindowTextLength(hWnd);
        if (length == 0) return string.Empty;
        StringBuilder sb = new StringBuilder(length + 1);
        GetWindowText(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }

    public static string GetWindowClassName(IntPtr hWnd)
    {
        StringBuilder sb = new StringBuilder(256);
        GetClassName(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }

    public static List<ChildControlInfo> GetChildControls(IntPtr hWndParent)
    {
        var children = new List<ChildControlInfo>();
        EnumChildWindows(hWndParent, (hWnd, lParam) =>
        {
            var info = new ChildControlInfo
            {
                Handle = hWnd,
                Text = GetWindowTitle(hWnd),
                ClassName = GetWindowClassName(hWnd)
            };
            children.Add(info);
            return true;
        }, IntPtr.Zero);
        return children;
    }

    public static List<DialogInfo> FindDialogsForProcess(uint pid)
    {
        var dialogs = new List<DialogInfo>();

        EnumWindows((hWnd, lParam) =>
        {
            if (!IsWindowVisible(hWnd)) return true;

            uint windowPid;
            GetWindowThreadProcessId(hWnd, out windowPid);
            if (windowPid != pid) return true;

            string className = GetWindowClassName(hWnd);

            // Skip the main Access window
            if (className == "OMain") return true;

            bool isDialogClass = className == "#32770";
            IntPtr owner = GetWindow(hWnd, GW_OWNER);
            bool isOwned = owner != IntPtr.Zero;
            int style = GetWindowLong(hWnd, GWL_STYLE);
            int exStyle = GetWindowLong(hWnd, GWL_EXSTYLE);
            bool hasModalFrame = (exStyle & WS_EX_DLGMODALFRAME) != 0;

            bool isDialog = isDialogClass || (isOwned && hasModalFrame);

            if (!isDialog) return true;

            var children = GetChildControls(hWnd);
            var childTexts = new List<string>();
            foreach (var child in children)
            {
                if (!string.IsNullOrWhiteSpace(child.Text))
                    childTexts.Add(child.Text);
            }

            dialogs.Add(new DialogInfo
            {
                Handle = hWnd,
                Title = GetWindowTitle(hWnd),
                ClassName = className,
                ProcessId = windowPid,
                IsDialog = true,
                ChildTexts = childTexts,
                ChildControls = children
            });

            return true;
        }, IntPtr.Zero);

        return dialogs;
    }

    public static bool DismissViaClose(IntPtr hWnd)
    {
        return PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
    }

    public static bool ClickButton(IntPtr hWndButton)
    {
        SendMessage(hWndButton, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
        return true;
    }

    public static IntPtr FindChildByText(IntPtr hWndParent, string text)
    {
        IntPtr found = IntPtr.Zero;
        EnumChildWindows(hWndParent, (hWnd, lParam) =>
        {
            string childText = GetWindowTitle(hWnd);
            if (string.Equals(childText, text, StringComparison.OrdinalIgnoreCase))
            {
                found = hWnd;
                return false; // stop enumerating
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    public static IntPtr FindChildByTextContains(IntPtr hWndParent, string substring)
    {
        IntPtr found = IntPtr.Zero;
        EnumChildWindows(hWndParent, (hWnd, lParam) =>
        {
            string childText = GetWindowTitle(hWnd);
            if (childText != null && childText.IndexOf(substring, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                found = hWnd;
                return false;
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }
}
'@ -ErrorAction SilentlyContinue

# ── Screenshot capture ─────────────────────────────────────────────────────────

function Invoke-ScreenshotCapture {
    param(
        [Parameter(Mandatory)]
        [string]$OutputPath
    )

    try {
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop

        $bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
        $bitmap = New-Object System.Drawing.Bitmap($bounds.Width, $bounds.Height)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)
        $graphics.Dispose()
        $bitmap.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
        $bitmap.Dispose()
        return $true
    }
    catch {
        Write-Host ("DIALOG_WATCHER: screenshot failed: {0}" -f $_.Exception.Message)
        return $false
    }
}

# ── Auto-dismiss logic ─────────────────────────────────────────────────────────

function Try-AutoDismissDialog {
    param(
        [Parameter(Mandatory)]
        [object]$Dialog
    )

    $title = [string]$Dialog.Title
    $childTextsJoined = ($Dialog.ChildTexts -join " ").ToLowerInvariant()
    $titleLower = $title.ToLowerInvariant()
    $handle = $Dialog.Handle

    # Patterns that should NOT be dismissed (log only)
    if ($childTextsJoined -match 'security warning|blocked|trust center') {
        return [PSCustomObject]@{ Dismissed = $false; Pattern = "security_warning"; Safe = $false; Reason = "Security/Trust Center dialog - needs config fix" }
    }
    if ($titleLower -match 'server is busy|com rpc|call was rejected') {
        return [PSCustomObject]@{ Dismissed = $false; Pattern = "server_busy_com_rpc"; Safe = $false; Reason = "COM RPC / server busy - indicates deadlock" }
    }

    # "Enter Parameter Value" — dismiss via WM_CLOSE (Escape)
    if ($titleLower -match 'enter parameter value') {
        [DialogDetector]::DismissViaClose($handle) | Out-Null
        return [PSCustomObject]@{ Dismissed = $true; Pattern = "enter_parameter_value"; Safe = $true; Reason = "Parameter prompt dismissed via WM_CLOSE" }
    }

    # "Do you want to save" — click "No"
    if ($childTextsJoined -match 'do you want to save') {
        $noButton = [DialogDetector]::FindChildByText($handle, "No")
        if ($noButton -eq [IntPtr]::Zero) {
            $noButton = [DialogDetector]::FindChildByText($handle, "&No")
        }
        if ($noButton -ne [IntPtr]::Zero) {
            [DialogDetector]::ClickButton($noButton) | Out-Null
            return [PSCustomObject]@{ Dismissed = $true; Pattern = "save_prompt"; Safe = $true; Reason = "Save prompt dismissed via No button" }
        }
        [DialogDetector]::DismissViaClose($handle) | Out-Null
        return [PSCustomObject]@{ Dismissed = $true; Pattern = "save_prompt"; Safe = $true; Reason = "Save prompt dismissed via WM_CLOSE (No button not found)" }
    }

    # VBA compile error
    if ($titleLower -match 'compile error' -or ($titleLower -match 'microsoft visual basic' -and $childTextsJoined -match 'compile error')) {
        [DialogDetector]::DismissViaClose($handle) | Out-Null
        return [PSCustomObject]@{ Dismissed = $true; Pattern = "vba_compile_error"; Safe = $true; Reason = "VBA compile error dismissed via WM_CLOSE" }
    }

    # "Action Failed"
    if ($titleLower -match 'action failed') {
        $stopButton = [DialogDetector]::FindChildByText($handle, "Stop")
        if ($stopButton -ne [IntPtr]::Zero) {
            [DialogDetector]::ClickButton($stopButton) | Out-Null
            return [PSCustomObject]@{ Dismissed = $true; Pattern = "action_failed"; Safe = $true; Reason = "Action Failed dismissed via Stop button" }
        }
        [DialogDetector]::DismissViaClose($handle) | Out-Null
        return [PSCustomObject]@{ Dismissed = $true; Pattern = "action_failed"; Safe = $true; Reason = "Action Failed dismissed via WM_CLOSE" }
    }

    # "can't find" / "not a valid path" / "file not found"
    if ($childTextsJoined -match "can't find|cannot find|not a valid path|could not find|file not found") {
        $okButton = [DialogDetector]::FindChildByText($handle, "OK")
        if ($okButton -ne [IntPtr]::Zero) {
            [DialogDetector]::ClickButton($okButton) | Out-Null
            return [PSCustomObject]@{ Dismissed = $true; Pattern = "not_found_path_error"; Safe = $true; Reason = "Path error dismissed via OK button" }
        }
        [DialogDetector]::DismissViaClose($handle) | Out-Null
        return [PSCustomObject]@{ Dismissed = $true; Pattern = "not_found_path_error"; Safe = $true; Reason = "Path error dismissed via WM_CLOSE" }
    }

    # Macro warning (SetWarnings)
    if ($titleLower -match '^action$' -and $childTextsJoined -match 'warning') {
        $yesButton = [DialogDetector]::FindChildByText($handle, "Yes")
        if ($yesButton -eq [IntPtr]::Zero) {
            $yesButton = [DialogDetector]::FindChildByText($handle, "&Yes")
        }
        if ($yesButton -ne [IntPtr]::Zero) {
            [DialogDetector]::ClickButton($yesButton) | Out-Null
            return [PSCustomObject]@{ Dismissed = $true; Pattern = "macro_warning"; Safe = $true; Reason = "Macro warning dismissed via Yes button" }
        }
    }

    # Generic Access/Office/VBA error dialogs with OK button
    if ($titleLower -match 'microsoft access|microsoft office|microsoft visual basic') {
        $okButton = [DialogDetector]::FindChildByText($handle, "OK")
        if ($okButton -ne [IntPtr]::Zero) {
            [DialogDetector]::ClickButton($okButton) | Out-Null
            return [PSCustomObject]@{ Dismissed = $true; Pattern = "generic_access_error"; Safe = $true; Reason = "Generic Access dialog dismissed via OK button" }
        }
    }

    # "Save As" dialog - dismiss via Escape (WM_CLOSE)
    if ($titleLower -match 'save as') {
        [DialogDetector]::DismissViaClose($handle) | Out-Null
        return [PSCustomObject]@{ Dismissed = $true; Pattern = "save_as_prompt"; Safe = $true; Reason = "Save As dialog dismissed via WM_CLOSE" }
    }

    # Unrecognized dialog - do not dismiss
    return [PSCustomObject]@{ Dismissed = $false; Pattern = "unknown"; Safe = $false; Reason = "Unrecognized dialog - not auto-dismissed" }
}

# ── Dialog watcher (background job) ───────────────────────────────────────────

function Start-DialogWatcher {
    param(
        [Parameter(Mandatory)]
        [string]$DiagnosticsPath,
        [string]$ScreenshotDir = $null,
        [int]$PollIntervalMs = 500,
        [switch]$AutoDismiss
    )

    if ([string]::IsNullOrWhiteSpace($ScreenshotDir)) {
        $ScreenshotDir = $DiagnosticsPath
    }

    if (-not (Test-Path $DiagnosticsPath)) {
        New-Item -ItemType Directory -Path $DiagnosticsPath -Force | Out-Null
    }
    if (-not (Test-Path $ScreenshotDir)) {
        New-Item -ItemType Directory -Path $ScreenshotDir -Force | Out-Null
    }

    $jsonlPath = Join-Path $DiagnosticsPath "dialog_events.jsonl"
    $autoDismissFlag = [bool]$AutoDismiss

    $job = Start-Job -ScriptBlock {
        param($JsonlPath, $ScreenshotDir, $PollIntervalMs, $AutoDismissFlag, $WatcherScriptPath)

        # Re-import the types inside the job
        . $WatcherScriptPath

        $seenDialogs = @{}

        while ($true) {
            try {
                $msAccessProcesses = @(Get-Process -Name MSACCESS -ErrorAction SilentlyContinue)
                foreach ($proc in $msAccessProcesses) {
                    $procId = [uint32]$proc.Id
                    $dialogs = [DialogDetector]::FindDialogsForProcess($procId)

                    foreach ($dialog in $dialogs) {
                        $handleKey = $dialog.Handle.ToString()
                        if ($seenDialogs.ContainsKey($handleKey)) {
                            continue
                        }
                        $seenDialogs[$handleKey] = $true

                        $timestamp = (Get-Date).ToUniversalTime().ToString("o")
                        $screenshotName = "dialog_{0}_pid{1}.png" -f ($timestamp -replace '[:\.]', ''), $procId
                        $screenshotPath = Join-Path $ScreenshotDir $screenshotName

                        # Take screenshot before any dismiss attempt
                        $screenshotOk = $false
                        try {
                            Add-Type -AssemblyName System.Drawing -ErrorAction Stop
                            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
                            $bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
                            $bitmap = New-Object System.Drawing.Bitmap($bounds.Width, $bounds.Height)
                            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
                            $graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)
                            $graphics.Dispose()
                            $bitmap.Save($screenshotPath, [System.Drawing.Imaging.ImageFormat]::Png)
                            $bitmap.Dispose()
                            $screenshotOk = $true
                        }
                        catch {
                            $screenshotPath = $null
                        }

                        $autoDismissed = $false
                        $pattern = "unknown"
                        $dismissReason = ""

                        if ($AutoDismissFlag) {
                            $result = Try-AutoDismissDialog -Dialog $dialog
                            $autoDismissed = [bool]$result.Dismissed
                            $pattern = [string]$result.Pattern
                            $dismissReason = [string]$result.Reason
                        }

                        $event = @{
                            timestamp = $timestamp
                            pid = [int]$procId
                            title = [string]$dialog.Title
                            className = [string]$dialog.ClassName
                            childTexts = @($dialog.ChildTexts)
                            screenshotPath = if ($screenshotOk) { $screenshotName } else { $null }
                            autoDismissed = $autoDismissed
                            pattern = $pattern
                            dismissReason = $dismissReason
                        }

                        $jsonLine = $event | ConvertTo-Json -Depth 10 -Compress
                        [System.IO.File]::AppendAllText($JsonlPath, $jsonLine + "`n")
                    }
                }
            }
            catch {
                # Swallow polling errors to keep the watcher alive
            }

            Start-Sleep -Milliseconds $PollIntervalMs
        }
    } -ArgumentList $jsonlPath, $ScreenshotDir, $PollIntervalMs, $autoDismissFlag, $script:_DialogWatcherScriptPath

    return [PSCustomObject]@{
        Job = $job
        DiagnosticsPath = $DiagnosticsPath
        JsonlPath = $jsonlPath
        ScreenshotDir = $ScreenshotDir
    }
}

function Stop-DialogWatcher {
    param(
        [Parameter(Mandatory)]
        [object]$WatcherState
    )

    if ($null -eq $WatcherState -or $null -eq $WatcherState.Job) {
        return
    }

    try {
        Stop-Job -Job $WatcherState.Job -ErrorAction SilentlyContinue
        Remove-Job -Job $WatcherState.Job -Force -ErrorAction SilentlyContinue
    }
    catch {
        # Ignore cleanup failures
    }
}

function Get-DialogWatcherReport {
    param(
        [Parameter(Mandatory)]
        [string]$JsonlPath
    )

    if (-not (Test-Path $JsonlPath)) {
        return @()
    }

    $events = @()
    foreach ($line in (Get-Content -Path $JsonlPath -ErrorAction SilentlyContinue)) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }
        try {
            $events += ($line | ConvertFrom-Json)
        }
        catch {
            # skip malformed lines
        }
    }

    return $events
}

function Write-DialogWatcherSummary {
    param(
        [Parameter(Mandatory)]
        [string]$JsonlPath,
        [string]$SectionName = ""
    )

    $events = Get-DialogWatcherReport -JsonlPath $JsonlPath
    if ($events.Count -eq 0) {
        return
    }

    $dismissed = @($events | Where-Object { $_.autoDismissed -eq $true }).Count
    $logged = @($events | Where-Object { $_.autoDismissed -ne $true }).Count

    $prefix = if (-not [string]::IsNullOrWhiteSpace($SectionName)) { "DIALOG_WATCHER[$SectionName]" } else { "DIALOG_WATCHER" }
    Write-Host ("{0}: {1} dialog(s) detected, {2} auto-dismissed, {3} logged-only" -f $prefix, $events.Count, $dismissed, $logged)

    foreach ($evt in $events) {
        $status = if ($evt.autoDismissed) { "DISMISSED" } else { "DETECTED" }
        Write-Host ("  {0} pid={1} title='{2}' pattern={3}" -f $status, $evt.pid, $evt.title, $evt.pattern)
    }
}

# ── Timeout-aware MCP batch execution ─────────────────────────────────────────

function Invoke-McpBatchWithTimeout {
    param(
        [Parameter(Mandatory)]
        [string]$ExePath,
        [Parameter(Mandatory)]
        [System.Collections.Generic.List[object]]$Calls,
        [string]$ClientName = "full-regression",
        [string]$ClientVersion = "1.0",
        [int]$TimeoutSeconds = 120,
        [string]$SectionName = "",
        [string]$ScreenshotDir = $null
    )

    $jsonLines = New-Object 'System.Collections.Generic.List[string]'
    $jsonLines.Add((@{
        jsonrpc = "2.0"
        id = 1
        method = "initialize"
        params = @{
            protocolVersion = "2024-11-05"
            capabilities = @{}
            clientInfo = @{
                name = $ClientName
                version = $ClientVersion
            }
        }
    } | ConvertTo-Json -Depth 40 -Compress))

    foreach ($call in $Calls) {
        $jsonLines.Add((@{
            jsonrpc = "2.0"
            id = $call.Id
            method = "tools/call"
            params = @{
                name = $call.Name
                arguments = $call.Arguments
            }
        } | ConvertTo-Json -Depth 50 -Compress))
    }

    $inputPayload = $jsonLines -join "`n"

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $ExePath
    $psi.UseShellExecute = $false
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi

    $stdoutBuilder = New-Object System.Text.StringBuilder
    $stderrBuilder = New-Object System.Text.StringBuilder

    $stdoutEvent = Register-ObjectEvent -InputObject $process -EventName OutputDataReceived -Action {
        if ($null -ne $EventArgs.Data) {
            [void]$Event.MessageData.AppendLine($EventArgs.Data)
        }
    } -MessageData $stdoutBuilder

    $stderrEvent = Register-ObjectEvent -InputObject $process -EventName ErrorDataReceived -Action {
        if ($null -ne $EventArgs.Data) {
            [void]$Event.MessageData.AppendLine($EventArgs.Data)
        }
    } -MessageData $stderrBuilder

    try {
        $process.Start() | Out-Null
        $process.BeginOutputReadLine()
        $process.BeginErrorReadLine()

        $process.StandardInput.Write($inputPayload)
        $process.StandardInput.Close()

        $exited = $process.WaitForExit($TimeoutSeconds * 1000)

        if (-not $exited) {
            # Timeout — capture diagnostics
            Write-Host ("BATCH_TIMEOUT: section='{0}' after {1}s" -f $SectionName, $TimeoutSeconds)

            if (-not [string]::IsNullOrWhiteSpace($ScreenshotDir)) {
                $tsName = (Get-Date).ToUniversalTime().ToString("yyyyMMddTHHmmss") + "Z"
                $timeoutScreenshot = Join-Path $ScreenshotDir ("timeout_{0}_{1}.png" -f ($SectionName -replace '[^a-zA-Z0-9_-]', '_'), $tsName)
                Invoke-ScreenshotCapture -OutputPath $timeoutScreenshot | Out-Null
            }

            try {
                $process.Kill()
                $process.WaitForExit(5000) | Out-Null
            }
            catch {
                # Process may have already exited
            }

            return @{ _timeout = $true; _section = $SectionName; _timeoutSeconds = $TimeoutSeconds }
        }

        # Ensure async reads complete
        $process.WaitForExit()
    }
    finally {
        Unregister-Event -SourceIdentifier $stdoutEvent.Name -ErrorAction SilentlyContinue
        Unregister-Event -SourceIdentifier $stderrEvent.Name -ErrorAction SilentlyContinue
        Remove-Job -Name $stdoutEvent.Name -Force -ErrorAction SilentlyContinue
        Remove-Job -Name $stderrEvent.Name -Force -ErrorAction SilentlyContinue
    }

    $rawOutput = $stdoutBuilder.ToString()
    $rawLines = $rawOutput -split "`r?`n"

    $responses = @{}
    foreach ($line in $rawLines) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        try {
            $parsed = $line | ConvertFrom-Json
            if ($null -ne $parsed.id) {
                $responses[[int]$parsed.id] = $parsed
            }
        }
        catch {
            Write-Host ("WARN: Could not parse response line: {0}" -f $line)
        }
    }

    return $responses
}

function Invoke-McpRawBatchWithTimeout {
    param(
        [Parameter(Mandatory)]
        [string]$ExePath,
        [Parameter(Mandatory)]
        [System.Collections.Generic.List[hashtable]]$Requests,
        [string]$ClientName = "full-regression-raw",
        [string]$ClientVersion = "1.0",
        [int]$TimeoutSeconds = 120,
        [string]$SectionName = "",
        [string]$ScreenshotDir = $null
    )

    $jsonLines = New-Object 'System.Collections.Generic.List[string]'
    $jsonLines.Add((@{
        jsonrpc = "2.0"
        id = 1
        method = "initialize"
        params = @{
            protocolVersion = "2024-11-05"
            capabilities = @{}
            clientInfo = @{
                name = $ClientName
                version = $ClientVersion
            }
        }
    } | ConvertTo-Json -Depth 40 -Compress))

    foreach ($req in $Requests) {
        $jsonLines.Add(($req | ConvertTo-Json -Depth 50 -Compress))
    }

    $inputPayload = $jsonLines -join "`n"

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $ExePath
    $psi.UseShellExecute = $false
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi

    $stdoutBuilder = New-Object System.Text.StringBuilder
    $stderrBuilder = New-Object System.Text.StringBuilder

    $stdoutEvent = Register-ObjectEvent -InputObject $process -EventName OutputDataReceived -Action {
        if ($null -ne $EventArgs.Data) {
            [void]$Event.MessageData.AppendLine($EventArgs.Data)
        }
    } -MessageData $stdoutBuilder

    $stderrEvent = Register-ObjectEvent -InputObject $process -EventName ErrorDataReceived -Action {
        if ($null -ne $EventArgs.Data) {
            [void]$Event.MessageData.AppendLine($EventArgs.Data)
        }
    } -MessageData $stderrBuilder

    try {
        $process.Start() | Out-Null
        $process.BeginOutputReadLine()
        $process.BeginErrorReadLine()

        $process.StandardInput.Write($inputPayload)
        $process.StandardInput.Close()

        $exited = $process.WaitForExit($TimeoutSeconds * 1000)

        if (-not $exited) {
            Write-Host ("BATCH_TIMEOUT: section='{0}' after {1}s" -f $SectionName, $TimeoutSeconds)

            if (-not [string]::IsNullOrWhiteSpace($ScreenshotDir)) {
                $tsName = (Get-Date).ToUniversalTime().ToString("yyyyMMddTHHmmss") + "Z"
                $timeoutScreenshot = Join-Path $ScreenshotDir ("timeout_{0}_{1}.png" -f ($SectionName -replace '[^a-zA-Z0-9_-]', '_'), $tsName)
                Invoke-ScreenshotCapture -OutputPath $timeoutScreenshot | Out-Null
            }

            try {
                $process.Kill()
                $process.WaitForExit(5000) | Out-Null
            }
            catch {}

            return @{ _timeout = $true; _section = $SectionName; _timeoutSeconds = $TimeoutSeconds }
        }

        $process.WaitForExit()
    }
    finally {
        Unregister-Event -SourceIdentifier $stdoutEvent.Name -ErrorAction SilentlyContinue
        Unregister-Event -SourceIdentifier $stderrEvent.Name -ErrorAction SilentlyContinue
        Remove-Job -Name $stdoutEvent.Name -Force -ErrorAction SilentlyContinue
        Remove-Job -Name $stderrEvent.Name -Force -ErrorAction SilentlyContinue
    }

    $rawOutput = $stdoutBuilder.ToString()
    $rawLines = $rawOutput -split "`r?`n"

    $responses = @{}
    foreach ($line in $rawLines) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        try {
            $parsed = $line | ConvertFrom-Json
            if ($null -ne $parsed.id) {
                $responses[[int]$parsed.id] = $parsed
            }
        }
        catch {
            Write-Host ("WARN: Could not parse response line: {0}" -f $line)
        }
    }

    return $responses
}

function Get-McpToolsListWithTimeout {
    param(
        [Parameter(Mandatory)]
        [string]$ExePath,
        [string]$ClientName = "full-regression-tools-list",
        [string]$ClientVersion = "1.0",
        [int]$TimeoutSeconds = 60,
        [string]$ScreenshotDir = $null
    )

    $calls = New-Object 'System.Collections.Generic.List[object]'
    # We use a dummy call list — the tools/list request is built inline
    $jsonLines = New-Object 'System.Collections.Generic.List[string]'
    $jsonLines.Add((@{
        jsonrpc = "2.0"
        id = 1
        method = "initialize"
        params = @{
            protocolVersion = "2024-11-05"
            capabilities = @{}
            clientInfo = @{
                name = $ClientName
                version = $ClientVersion
            }
        }
    } | ConvertTo-Json -Depth 40 -Compress))

    $jsonLines.Add((@{
        jsonrpc = "2.0"
        id = 2
        method = "tools/list"
        params = @{}
    } | ConvertTo-Json -Depth 40 -Compress))

    $inputPayload = $jsonLines -join "`n"

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $ExePath
    $psi.UseShellExecute = $false
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi

    $stdoutBuilder = New-Object System.Text.StringBuilder
    $stderrBuilder = New-Object System.Text.StringBuilder

    $stdoutEvent = Register-ObjectEvent -InputObject $process -EventName OutputDataReceived -Action {
        if ($null -ne $EventArgs.Data) {
            [void]$Event.MessageData.AppendLine($EventArgs.Data)
        }
    } -MessageData $stdoutBuilder

    $stderrEvent = Register-ObjectEvent -InputObject $process -EventName ErrorDataReceived -Action {
        if ($null -ne $EventArgs.Data) {
            [void]$Event.MessageData.AppendLine($EventArgs.Data)
        }
    } -MessageData $stderrBuilder

    try {
        $process.Start() | Out-Null
        $process.BeginOutputReadLine()
        $process.BeginErrorReadLine()

        $process.StandardInput.Write($inputPayload)
        $process.StandardInput.Close()

        $exited = $process.WaitForExit($TimeoutSeconds * 1000)

        if (-not $exited) {
            Write-Host ("TOOLS_LIST_TIMEOUT: after {0}s" -f $TimeoutSeconds)

            if (-not [string]::IsNullOrWhiteSpace($ScreenshotDir)) {
                $tsName = (Get-Date).ToUniversalTime().ToString("yyyyMMddTHHmmss") + "Z"
                $timeoutScreenshot = Join-Path $ScreenshotDir ("timeout_tools_list_{0}.png" -f $tsName)
                Invoke-ScreenshotCapture -OutputPath $timeoutScreenshot | Out-Null
            }

            try {
                $process.Kill()
                $process.WaitForExit(5000) | Out-Null
            }
            catch {}

            return @()
        }

        $process.WaitForExit()
    }
    finally {
        Unregister-Event -SourceIdentifier $stdoutEvent.Name -ErrorAction SilentlyContinue
        Unregister-Event -SourceIdentifier $stderrEvent.Name -ErrorAction SilentlyContinue
        Remove-Job -Name $stdoutEvent.Name -Force -ErrorAction SilentlyContinue
        Remove-Job -Name $stderrEvent.Name -Force -ErrorAction SilentlyContinue
    }

    $rawOutput = $stdoutBuilder.ToString()
    $rawLines = $rawOutput -split "`r?`n"

    $responses = @{}
    foreach ($line in $rawLines) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        try {
            $parsed = $line | ConvertFrom-Json
            if ($null -ne $parsed.id) {
                $responses[[int]$parsed.id] = $parsed
            }
        }
        catch {
            Write-Host ("WARN: Could not parse tools/list response line: {0}" -f $line)
        }
    }

    if (-not $responses.ContainsKey(2)) {
        return @()
    }

    $toolResponse = $responses[2]
    if ($toolResponse.result -and $toolResponse.result.tools) {
        return @($toolResponse.result.tools)
    }

    return @()
}

# ── Diagnostics summary ──────────────────────────────────────────────────────

function Write-DiagnosticsSummary {
    param(
        [Parameter(Mandatory)]
        [string]$DiagnosticsPath,
        [string]$JsonlPath = $null,
        [int]$TotalFailed = 0,
        [int]$TimeoutCount = 0,
        [hashtable]$TimeoutSections = @{}
    )

    if ([string]::IsNullOrWhiteSpace($JsonlPath)) {
        $JsonlPath = Join-Path $DiagnosticsPath "dialog_events.jsonl"
    }

    $events = Get-DialogWatcherReport -JsonlPath $JsonlPath
    $dismissed = @($events | Where-Object { $_.autoDismissed -eq $true }).Count
    $loggedOnly = @($events | Where-Object { $_.autoDismissed -ne $true }).Count

    $summary = @{
        timestamp = (Get-Date).ToUniversalTime().ToString("o")
        totalDialogEvents = $events.Count
        autoDismissed = $dismissed
        loggedOnly = $loggedOnly
        totalFailed = $TotalFailed
        timeoutCount = $TimeoutCount
        timeoutSections = @($TimeoutSections.Keys)
        events = @($events)
    }

    $summaryPath = Join-Path $DiagnosticsPath "summary.json"
    $summary | ConvertTo-Json -Depth 20 | Set-Content -Path $summaryPath -Encoding UTF8

    Write-Host ("DIAGNOSTICS_SUMMARY: {0} dialog(s), {1} dismissed, {2} logged-only, {3} timeouts" -f $events.Count, $dismissed, $loggedOnly, $TimeoutCount)
}
