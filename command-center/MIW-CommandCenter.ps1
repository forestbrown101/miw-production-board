# ============================================================
#  MIW COMMAND CENTER
#  Opens a production board in a clean Chromium window, or the
#  latest weekly backlog workbook in Excel.
#  Everything configurable lives in boards.json next to this file.
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Give this window its OWN taskbar identity. The launcher runs inside powershell.exe, and
# Windows groups taskbar buttons by process, so without this the button shows the
# PowerShell icon even though the window itself carries MIW.ico. Registering an explicit
# AppUserModelID before any window is created makes Windows treat the launcher as its own
# app and use the form's icon (the MIW logo). Must run before the form exists.
$CCAppId = 'MerchantIronWorks.CommandCenter'   # same string is stamped onto shortcuts, see Set-ShortcutAppId
try {
    Add-Type -Namespace MIW -Name Shell -MemberDefinition @'
[DllImport("shell32.dll", SetLastError = true)]
public static extern int SetCurrentProcessExplicitAppUserModelID([MarshalAs(UnmanagedType.LPWStr)] string AppID);
'@
    [MIW.Shell]::SetCurrentProcessExplicitAppUserModelID($CCAppId) | Out-Null
} catch { }

# Where am I? Two ways this runs:
#   1. as the .ps1 (double-click the .bat / a shortcut) -> $PSScriptRoot is this script's folder
#   2. compiled into MIW Command Center.exe (PS2EXE, no console) -> $PSScriptRoot is empty, so
#      use the exe's own folder. The exe may sit at the folder ROOT with the support files in
#      App\, so if boards.json is not beside us but is in .\App, use App as the working folder.
$__proc  = [System.Diagnostics.Process]::GetCurrentProcess()
$ExePath = if ($__proc.ProcessName -notmatch '^(powershell|powershell_ise|pwsh)$') { $__proc.MainModule.FileName } else { $null }
$ScriptDir = if ($ExePath) { Split-Path -Parent $ExePath }
             elseif ($PSScriptRoot) { $PSScriptRoot }
             else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not (Test-Path (Join-Path $ScriptDir 'boards.json')) -and (Test-Path (Join-Path $ScriptDir 'App\boards.json'))) {
    $ScriptDir = Join-Path $ScriptDir 'App'
}
$ConfigPath = Join-Path $ScriptDir 'boards.json'
$IconPath   = Join-Path $ScriptDir 'MIW.ico'

# ---------- colors ----------
function C([string]$hex) { [System.Drawing.ColorTranslator]::FromHtml($hex) }
$BG      = C '#0E131D'
$CARD    = C '#161C29'
$CARDHOV = C '#1E2636'
$TEXT    = C '#E6EDF7'
$MUTED   = C '#8A97AD'
$DIM     = C '#4A5468'

# ---------- config ----------
if (-not (Test-Path $ConfigPath)) {
    [System.Windows.Forms.MessageBox]::Show("boards.json was not found next to this launcher.`r`n`r`nExpected:`r`n$ConfigPath", 'MIW Command Center', 'OK', 'Error') | Out-Null
    exit 1
}
try { $Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json }
catch {
    [System.Windows.Forms.MessageBox]::Show("boards.json could not be read. It may have a typo.`r`n`r`n$($_.Exception.Message)", 'MIW Command Center', 'OK', 'Error') | Out-Null
    exit 1
}

# ---------- layout, sized to the tile count ----------
$TileRows = [int][math]::Ceiling( @($Config.boards).Count / 2 )
$FlowH    = ($TileRows * 108) + 8
$FootY    = 104 + $FlowH + 8

# ---------- find OneDrive root ----------
function Get-OneDriveRoots {
    $roots = New-Object System.Collections.Generic.List[string]
    foreach ($v in @($env:OneDriveCommercial, $env:OneDrive, $env:OneDriveConsumer)) {
        if ($v -and (Test-Path $v)) { $roots.Add($v) }
    }
    if (Test-Path $env:USERPROFILE) {
        Get-ChildItem -Path $env:USERPROFILE -Directory -Filter 'OneDrive*' -ErrorAction SilentlyContinue |
            ForEach-Object { $roots.Add($_.FullName) }
    }
    return ($roots | Select-Object -Unique)
}

# ---------- turn a configured folder into a real folder on this PC ----------
function Resolve-Folder([string]$f) {
    if ([string]::IsNullOrWhiteSpace($f)) { return $null }
    $f = [System.Environment]::ExpandEnvironmentVariables($f)
    if ([System.IO.Path]::IsPathRooted($f)) {
        if (Test-Path $f) { return (Resolve-Path $f).Path }
        return $null
    }
    $local = Join-Path $ScriptDir $f
    if (Test-Path $local) { return (Resolve-Path $local).Path }
    foreach ($root in (Get-OneDriveRoots)) {
        $p = Join-Path $root $f
        if (Test-Path $p) { return (Resolve-Path $p).Path }
    }
    return $null
}

# ---------- BOARD: resolve to a launchable address ----------
# Tries, in order: a web address, an absolute path, beside the launcher,
# each OneDrive root, then a recursive search under OneDrive.
function Resolve-Board($b) {
    if ([string]::IsNullOrWhiteSpace($b.target)) { return $null }
    $t = [System.Environment]::ExpandEnvironmentVariables($b.target)

    if ($t -match '^https?://') { return $t }

    if ([System.IO.Path]::IsPathRooted($t) -and (Test-Path $t)) {
        return ([uri](Resolve-Path $t).Path).AbsoluteUri
    }

    $local = Join-Path $ScriptDir $t
    if (Test-Path $local) { return ([uri](Resolve-Path $local).Path).AbsoluteUri }

    foreach ($root in (Get-OneDriveRoots)) {
        $p = Join-Path $root $t
        if (Test-Path $p) { return ([uri](Resolve-Path $p).Path).AbsoluteUri }
    }

    $leaf = Split-Path $t -Leaf
    foreach ($root in (Get-OneDriveRoots)) {
        $hit = Get-ChildItem -Path $root -Filter $leaf -Recurse -File -ErrorAction SilentlyContinue |
               Where-Object { $_.FullName -notmatch '\\_archive\\' } |
               Select-Object -First 1
        if ($hit) { return ([uri]$hit.FullName).AbsoluteUri }
    }

    return $null
}

# ---------- EXCEL: find the newest matching workbook ----------
# Ranks by WEEK number, then REV number, then last modified. Skips anything
# whose name contains one of the exclude words, plus Excel lock files and
# hidden OneDrive copies.
function Find-LatestWorkbook($b) {
    $dir = Resolve-Folder $b.folder
    if (-not $dir) { return $null }

    $pattern = if ($b.pattern) { $b.pattern } else { '*.xlsx' }
    $exclude = if ($b.exclude) { @($b.exclude) } else { @() }

    $files = Get-ChildItem -LiteralPath $dir -Filter $pattern -File -ErrorAction SilentlyContinue |
             Where-Object {
                 $n = $_.Name
                 if ($n.StartsWith('~$') -or $n.StartsWith('.')) { return $false }
                 foreach ($x in $exclude) {
                     if (-not [string]::IsNullOrWhiteSpace($x) -and $n -like "*$x*") { return $false }
                 }
                 return $true
             }
    if (-not $files) { return $null }

    $ranked = $files | ForEach-Object {
        $wk = 0; $rev = 0
        if ($_.BaseName -match 'WEEK\s*0*(\d+)') { $wk  = [int]$Matches[1] }
        if ($_.BaseName -match 'REV\s*0*(\d+)')  { $rev = [int]$Matches[1] }
        [pscustomobject]@{ File = $_; Week = $wk; Rev = $rev; Modified = $_.LastWriteTime }
    }
    return ($ranked | Sort-Object Week, Rev, Modified -Descending | Select-Object -First 1).File
}

# ---------- find a Chromium browser ----------
function Get-Browser {
    $pref = if ($Config.browser) { $Config.browser } else { 'auto' }
    $pf   = ${env:ProgramFiles}
    $pf86 = ${env:ProgramFiles(x86)}
    $lad  = $env:LOCALAPPDATA

    $edge = @(
        (Join-Path $pf86 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $pf   'Microsoft\Edge\Application\msedge.exe')
    )
    $chrome = @(
        (Join-Path $pf   'Google\Chrome\Application\chrome.exe'),
        (Join-Path $pf86 'Google\Chrome\Application\chrome.exe'),
        (Join-Path $lad  'Google\Chrome\Application\chrome.exe')
    )
    $order = switch ($pref) {
        'chrome' { $chrome + $edge }
        'edge'   { $edge + $chrome }
        default  { $edge + $chrome }
    }
    foreach ($c in $order) { if ($c -and (Test-Path $c)) { return $c } }
    return $null
}

function Get-ExcelExe {
    try {
        $k = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe' -ErrorAction Stop
        $p = $k.'(default)'
        if ($p -and (Test-Path $p)) { return $p }
    } catch { }
    foreach ($p in @(
        (Join-Path ${env:ProgramFiles}      'Microsoft Office\root\Office16\EXCEL.EXE'),
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office\root\Office16\EXCEL.EXE')
    )) { if ($p -and (Test-Path $p)) { return $p } }
    return $null
}

# ---------- launch ----------
function Open-BoardWindow($b, [bool]$tvMode) {
    $addr = Resolve-Board $b
    if (-not $addr) {
        [System.Windows.Forms.MessageBox]::Show(
            "Could not find the file for `"$($b.name)`".`r`n`r`nCheck the target in boards.json, and make sure the OneDrive folder has finished syncing on this PC.",
            'MIW Command Center', 'OK', 'Warning') | Out-Null
        return $null
    }
    $exe = Get-Browser
    if (-not $exe) {
        [System.Windows.Forms.MessageBox]::Show(
            "No Chromium browser was found on this PC (looked for Microsoft Edge and Google Chrome).",
            'MIW Command Center', 'OK', 'Error') | Out-Null
        return $null
    }

    $profileDir = Join-Path $env:LOCALAPPDATA 'MIWBoards\browser'
    if (-not (Test-Path $profileDir)) { New-Item -ItemType Directory -Path $profileDir -Force | Out-Null }

    $argList = @(
        "--app=$addr",
        "--user-data-dir=$profileDir",
        '--no-first-run',
        '--no-default-browser-check',
        '--disable-features=Translate,DefaultBrowserPromptRefresh',
        '--disable-background-mode'
    )
    if ($tvMode) { $argList += '--start-fullscreen' } else { $argList += '--start-maximized' }

    Start-Process -FilePath $exe -ArgumentList $argList | Out-Null
    return $b.name
}

function Open-Workbook($b) {
    # PREFERRED: the server publishes the newest workbook's direct SharePoint link into
    # backlog_latest.json beside this launcher every 5 minutes (publish_backlog_link.py).
    # Opening that link needs NO locally synced 2026 folder on this PC; the ms-excel:
    # protocol hands it straight to desktop Excel as the live, co-authored workbook.
    $linkName = if ($b.link) { $b.link } else { 'backlog_latest.json' }
    $linkPath = Join-Path $ScriptDir $linkName
    if (Test-Path -LiteralPath $linkPath) {
        try {
            $j = Get-Content -LiteralPath $linkPath -Raw | ConvertFrom-Json
            if ($j.excel) {
                Start-Process $j.excel | Out-Null
                return $j.name
            }
        } catch { }
    }

    # FALLBACK 1: the original behaviour, newest matching workbook in the locally synced
    # folder, opened with desktop Excel.
    $file = Find-LatestWorkbook $b
    if ($file) {
        $excel = Get-ExcelExe
        if ($excel) { Start-Process -FilePath $excel -ArgumentList @("`"$($file.FullName)`"") | Out-Null }
        else        { Start-Process -FilePath $file.FullName | Out-Null }   # let Windows decide
        return $file.Name
    }

    # FALLBACK 2: open the SharePoint folder view in the browser so the person can click
    # the file themselves. Works for anyone with permission, no sync required.
    if ($b.folder_url) {
        Start-Process $b.folder_url | Out-Null
        return 'SharePoint folder'
    }

    [System.Windows.Forms.MessageBox]::Show(
        "No workbook matched for `"$($b.name)`".`r`n`r`nFolder: $($b.folder)`r`nPattern: $($b.pattern)`r`n`r`nCheck boards.json, and make sure that OneDrive folder has finished syncing on this PC.",
        'MIW Command Center', 'OK', 'Warning') | Out-Null
    return $null
}

# ---------- shortcuts ----------
# Shortcuts point at powershell.exe running the .ps1, NOT at the .bat. Windows
# will not pin a shortcut whose target is a batch file, and this way there is
# no console window flash either. The MIW icon is what people actually see.
$PsExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

# Stamp a .lnk with the SAME AppUserModelID the running launcher registers at startup.
# Windows matches a taskbar pin to a live window by that ID, so without it a pinned
# shortcut and the open launcher show as TWO buttons. With it they merge into one.
# The ID is written as the shortcut's System.AppUserModel.ID property via the shell's
# property store (the same mechanism Windows itself uses).
function Set-ShortcutAppId([string]$lnkPath, [string]$appId) {
    try {
        if (-not ('MIW.LnkAppId' -as [type])) {
            Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace MIW {
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct PROPERTYKEY { public Guid fmtid; public uint pid; }
    [StructLayout(LayoutKind.Sequential)]
    public struct PROPVARIANT { public ushort vt; public ushort r1; public ushort r2; public ushort r3; public IntPtr p; public IntPtr p2; }
    [ComImport, Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPropertyStore {
        // [PreserveSig] makes the int the real HRESULT return. Without it .NET would treat
        // the int as a hidden trailing [retval] parameter and misalign every call.
        [PreserveSig] int GetCount(out uint cProps);
        [PreserveSig] int GetAt(uint iProp, out PROPERTYKEY pkey);
        [PreserveSig] int GetValue(ref PROPERTYKEY key, out PROPVARIANT pv);
        [PreserveSig] int SetValue(ref PROPERTYKEY key, ref PROPVARIANT pv);
        [PreserveSig] int Commit();
    }
    public static class LnkAppId {
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        static extern void SHGetPropertyStoreFromParsingName(string pszPath, IntPtr pbc, int flags, ref Guid riid, out IPropertyStore store);
        [DllImport("propsys.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        static extern void InitPropVariantFromString(string psz, out PROPVARIANT ppropvar);
        [DllImport("ole32.dll", PreserveSig = false)]
        static extern void PropVariantClear(ref PROPVARIANT pvar);
        public static void Set(string lnkPath, string appId) {
            Guid iid = new Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99");
            IPropertyStore ps;
            SHGetPropertyStoreFromParsingName(lnkPath, IntPtr.Zero, 2, ref iid, out ps);   // 2 = GPS_READWRITE
            PROPERTYKEY key = new PROPERTYKEY();
            key.fmtid = new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3");                   // System.AppUserModel.ID
            key.pid = 5;
            PROPVARIANT pv;
            InitPropVariantFromString(appId, out pv);
            try {
                Marshal.ThrowExceptionForHR(ps.SetValue(ref key, ref pv));
                Marshal.ThrowExceptionForHR(ps.Commit());
            }
            finally { PropVariantClear(ref pv); Marshal.ReleaseComObject(ps); }
        }
    }
}
'@
        }
        [MIW.LnkAppId]::Set($lnkPath, $appId)
    } catch { }   # never let a stamping problem stop the shortcut from being created
}

function New-CCShortcut([string]$lnkPath) {
    $sh = New-Object -ComObject WScript.Shell
    $sc = $sh.CreateShortcut($lnkPath)
    # Prefer the no-console launcher exe at the folder ROOT (built from App\Launcher.cs) if
    # it exists, so shortcuts and pins open the menu with no PowerShell/Terminal window.
    $rootExe = Join-Path (Split-Path -Parent $ScriptDir) 'MIW Command Center.exe'
    if (-not $ExePath -and (Test-Path -LiteralPath $rootExe)) { $ExePath = $rootExe }
    if ($ExePath) {
        # Point the shortcut straight at the exe: no PowerShell, no console, and the pin
        # inherits the exe's own MIW icon.
        $sc.TargetPath       = $ExePath
        $sc.Arguments        = ''
        $sc.WorkingDirectory = Split-Path -Parent $ExePath
    } else {
        $sc.TargetPath       = $PsExe
        $sc.Arguments        = '-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "' + (Join-Path $ScriptDir 'MIW-CommandCenter.ps1') + '"'
        $sc.WorkingDirectory = $ScriptDir
    }
    $sc.Description      = 'MIW production boards and the current backlog'
    if (Test-Path $IconPath) { $sc.IconLocation = "$IconPath,0" }
    $sc.Save()
    Set-ShortcutAppId $lnkPath $CCAppId   # merge the pin with the live window on the taskbar
    return $lnkPath
}

# The taskbar pin verb is named in the user's own language, so read it out of
# shell32 rather than hardcoding English.
function Get-MuiString([string]$ref) {
    try {
        if (-not ('MIW.Mui' -as [type])) {
            Add-Type -Namespace MIW -Name Mui -MemberDefinition @'
[DllImport("shlwapi.dll", CharSet=CharSet.Unicode)]
public static extern int SHLoadIndirectString(string pszSource, System.Text.StringBuilder pszOutBuf, int cchOutBuf, System.IntPtr ppvReserved);
'@
        }
        $sb = New-Object System.Text.StringBuilder 512
        if ([MIW.Mui]::SHLoadIndirectString($ref, $sb, $sb.Capacity, [IntPtr]::Zero) -eq 0) {
            return $sb.ToString()
        }
    } catch { }
    return $null
}

# Returns $true only if Windows actually let us pin. Windows 11 and later
# builds of Windows 10 hide this verb from anything that is not Explorer, so a
# $false here is normal, not a bug. The caller falls back to showing the file.
function Invoke-TaskbarPin([string]$lnkPath) {
    try {
        $shell = New-Object -ComObject Shell.Application
        $item  = $shell.Namespace((Split-Path $lnkPath -Parent)).ParseName((Split-Path $lnkPath -Leaf))
        if (-not $item) { return $false }
        $wanted = @()
        $mui = Get-MuiString '@shell32.dll,-5386'
        if ($mui) { $wanted += $mui }
        $wanted += 'Pin to tas&kbar', 'Pin to taskbar'
        foreach ($v in $item.Verbs()) {
            foreach ($n in $wanted) {
                if ($v.Name.Replace('&','') -eq $n.Replace('&','')) { $v.DoIt(); return $true }
            }
        }
    } catch { }
    return $false
}

# ============================================================
#  UI
# ============================================================
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'MIW Command Center'
$form.ClientSize      = New-Object System.Drawing.Size(704, ($FootY + 74))
$form.StartPosition   = 'CenterScreen'
$form.BackColor       = $BG
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox     = $false
if (Test-Path $IconPath) {
    try { $form.Icon = New-Object System.Drawing.Icon($IconPath) } catch { }
}

$hdr = New-Object System.Windows.Forms.Panel
$hdr.Size = New-Object System.Drawing.Size(704, 92)
$hdr.Location = New-Object System.Drawing.Point(0,0)
$hdr.BackColor = $BG
$form.Controls.Add($hdr)

$title = New-Object System.Windows.Forms.Label
$title.Text = 'MIW COMMAND CENTER'
$title.Font = New-Object System.Drawing.Font('Segoe UI', 17, [System.Drawing.FontStyle]::Bold)
$title.ForeColor = $TEXT
$title.AutoSize = $true
$title.Location = New-Object System.Drawing.Point(28, 24)
$hdr.Controls.Add($title)

$sub = New-Object System.Windows.Forms.Label
$sub.Text = 'Click one. Boards open in their own window, the backlog opens in Excel.'
$sub.Font = New-Object System.Drawing.Font('Segoe UI', 9.5)
$sub.ForeColor = $MUTED
$sub.AutoSize = $true
$sub.Location = New-Object System.Drawing.Point(30, 58)
$hdr.Controls.Add($sub)

$rule = New-Object System.Windows.Forms.Panel
$rule.Size = New-Object System.Drawing.Size(660, 1)
$rule.Location = New-Object System.Drawing.Point(28, 88)
$rule.BackColor = C '#232C3D'
$hdr.Controls.Add($rule)

$flow = New-Object System.Windows.Forms.FlowLayoutPanel
$flow.Location = New-Object System.Drawing.Point(20, 104)
$flow.Size = New-Object System.Drawing.Size(672, $FlowH)
$flow.BackColor = $BG
$flow.AutoScroll = $true
$flow.FlowDirection = 'LeftToRight'
$flow.WrapContents = $true
$form.Controls.Add($flow)

$tv = New-Object System.Windows.Forms.CheckBox
$tv.Text = 'TV mode (full screen)'
$tv.ForeColor = $MUTED
$tv.BackColor = $BG
$tv.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$tv.AutoSize = $true
$tv.Location = New-Object System.Drawing.Point(30, ($FootY + 34))
$form.Controls.Add($tv)

$status = New-Object System.Windows.Forms.Label
$status.Text = ''
$status.ForeColor = $MUTED
$status.BackColor = $BG
$status.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
$status.AutoSize = $false
$status.AutoEllipsis = $true
$status.TextAlign = 'MiddleLeft'
$status.Size = New-Object System.Drawing.Size(650, 20)
$status.Location = New-Object System.Drawing.Point(30, $FootY)
$form.Controls.Add($status)

# NOTE: parameter names here are deliberately ugly. PowerShell variable names
# are CASE INSENSITIVE, so a parameter called $text would be the same variable
# as the $TEXT colour constant, and $b.ForeColor = $TEXT would quietly try to
# set the colour to the button's own caption and leave the text unreadable.
function New-FooterButton([string]$btnLabel, [int]$btnX, [int]$btnW) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $btnLabel
    $b.Size = New-Object System.Drawing.Size($btnW, 30)
    $b.Location = New-Object System.Drawing.Point($btnX, ($FootY + 28))
    $b.FlatStyle = 'Flat'
    $b.FlatAppearance.BorderColor = C '#2C3547'
    $b.BackColor = $CARD
    $b.ForeColor = $TEXT
    $b.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $b.UseVisualStyleBackColor = $false
    return $b
}

$btnTask = New-FooterButton 'Pin to taskbar' 284 130
$btnTask.Add_Click({
    try {
        $startDir = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs'
        if (-not (Test-Path $startDir)) { New-Item -ItemType Directory -Path $startDir -Force | Out-Null }
        $lnk = New-CCShortcut (Join-Path $startDir 'MIW Command Center.lnk')

        if (Invoke-TaskbarPin $lnk) {
            $status.Text = 'Pinned to your taskbar.'
        } else {
            Start-Process explorer.exe "/select,`"$lnk`""
            $status.Text = 'Windows will not let an app pin itself. Right-click the highlighted shortcut, Show more options, Pin to taskbar.'
        }
    } catch {
        $status.Text = 'Could not pin: ' + $_.Exception.Message
    }
})
$form.Controls.Add($btnTask)

$btnPin = New-FooterButton 'Add desktop shortcut' 422 160
$btnPin.Add_Click({
    try {
        $lnk = New-CCShortcut (Join-Path ([Environment]::GetFolderPath('Desktop')) 'MIW Command Center.lnk')
        $status.Text = 'Shortcut added to your desktop.'
    } catch {
        $status.Text = 'Could not create the shortcut: ' + $_.Exception.Message
    }
})
$form.Controls.Add($btnPin)

$btnClose = New-FooterButton 'Close' 590 96
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

# ---------- build tiles ----------
foreach ($b in $Config.boards) {

    $isLive = ($b.status -ne 'soon')
    $accent = if ($b.color) { C $b.color } else { C '#5AA9FF' }

    # For an Excel card, show the file it actually found as the subtitle.
    $subText = $b.subtitle
    if ($isLive -and $b.type -eq 'excel') {
        $found = Find-LatestWorkbook $b
        if ($found) { $subText = $found.BaseName } else { $subText = "$($b.subtitle)   .   NOT FOUND ON THIS PC" }
    }
    if (-not $isLive) { $subText = "$($b.subtitle)   .   NOT WIRED UP YET" }

    $tile = New-Object System.Windows.Forms.Panel
    $tile.Size = New-Object System.Drawing.Size(312, 92)
    $tile.Margin = New-Object System.Windows.Forms.Padding(8)
    $tile.BackColor = $CARD
    $tile.Cursor = 'Hand'
    $tile.Tag = $b

    $bar = New-Object System.Windows.Forms.Panel
    $bar.Size = New-Object System.Drawing.Size(5, 92)
    $bar.Location = New-Object System.Drawing.Point(0,0)
    $bar.BackColor = if ($isLive) { $accent } else { $DIM }
    $tile.Controls.Add($bar)

    $lname = New-Object System.Windows.Forms.Label
    $lname.Text = $b.name
    $lname.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
    $lname.ForeColor = if ($isLive) { $TEXT } else { $DIM }
    $lname.AutoSize = $true
    $lname.MaximumSize = New-Object System.Drawing.Size(280, 0)
    $lname.Location = New-Object System.Drawing.Point(20, 16)
    $lname.BackColor = [System.Drawing.Color]::Transparent
    $lname.Cursor = 'Hand'
    $tile.Controls.Add($lname)

    $lsub = New-Object System.Windows.Forms.Label
    $lsub.Text = $subText
    $lsub.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $lsub.ForeColor = if ($isLive) { $MUTED } else { $DIM }
    $lsub.AutoSize = $true
    $lsub.MaximumSize = New-Object System.Drawing.Size(282, 34)
    $lsub.Location = New-Object System.Drawing.Point(22, 44)
    $lsub.BackColor = [System.Drawing.Color]::Transparent
    $lsub.Cursor = 'Hand'
    $tile.Controls.Add($lsub)

    $handler = {
        param($sender, $e)
        $panel = $sender
        while ($panel -and -not $panel.Tag) { $panel = $panel.Parent }
        if (-not $panel) { return }
        $bd = $panel.Tag

        if ($bd.status -eq 'soon') {
            [System.Windows.Forms.MessageBox]::Show(
                "$($bd.name) is not wired up yet.`r`n`r`nWhen it is ready, add its file to boards.json and the tile turns on here.",
                'MIW Command Center', 'OK', 'Information') | Out-Null
            return
        }

        $status.Text = "Opening $($bd.name) ..."
        $form.Refresh()

        if ($bd.type -eq 'excel') {
            $opened = Open-Workbook $bd
            $status.Text = if ($opened) { "Opened $opened" } else { "Could not open $($bd.name)." }
        } else {
            $opened = Open-BoardWindow $bd $tv.Checked
            $status.Text = if ($opened) { "$opened opened." } else { "Could not open $($bd.name)." }
        }
    }

    $enter = { param($s,$e) $p=$s; while($p -and -not $p.Tag){$p=$p.Parent}; if($p){$p.BackColor=$CARDHOV} }
    $leave = { param($s,$e) $p=$s; while($p -and -not $p.Tag){$p=$p.Parent}; if($p){$p.BackColor=$CARD} }

    foreach ($c in @($tile, $bar, $lname, $lsub)) {
        $c.Add_Click($handler)
        $c.Add_MouseEnter($enter)
        $c.Add_MouseLeave($leave)
    }

    $flow.Controls.Add($tile)
}

# The folder shortcut launches PowerShell MINIMIZED so the console stays out of the way.
# On Windows 11 (Windows Terminal) that "minimized" instruction also lands on the first
# window this process creates, which is this menu. So the moment the form is shown,
# force it back to a normal window and bring it to the front. Net effect, by design:
# PowerShell opens minimized, the Command Center pops up on top.
$form.WindowState = 'Normal'
$form.Add_Shown({
    $form.WindowState = 'Normal'
    $form.TopMost = $true          # briefly, to win focus from the launching window
    $form.Activate()
    $form.TopMost = $false
})

[void]$form.ShowDialog()
