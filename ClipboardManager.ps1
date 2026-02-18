<#
.SYNOPSIS
    Clipboard Manager - Stealth Edition (PowerShell)
.DESCRIPTION
    Lightweight clipboard history manager with image support.
    No external dependencies - runs on any Windows 10/11 machine.
    Hotkey: Ctrl+Shift+H = Toggle window
#>

# Self-relaunch as fully hidden process (no console, no taskbar)
if ($args -notcontains '-Hidden') {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "powershell.exe"
    $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`" -Hidden"
    $psi.CreateNoWindow = $true
    $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
    $psi.UseShellExecute = $false
    [System.Diagnostics.Process]::Start($psi) | Out-Null
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- P/Invoke for global hotkey and stealth ---
Add-Type -ReferencedAssemblies System.Windows.Forms, System.Drawing -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class NativeMethods {
    [DllImport("user32.dll")]
    public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll")]
    public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    [DllImport("user32.dll")]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);
    [DllImport("user32.dll")]
    public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
    [DllImport("user32.dll")]
    public static extern bool ReleaseCapture();
    [DllImport("user32.dll")]
    public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

    public const int GWL_EXSTYLE = -20;
    public const int WS_EX_TOOLWINDOW = 0x80;
    public const int WM_NCLBUTTONDOWN = 0xA1;
    public const int HT_CAPTION = 0x2;
    public const int WM_HOTKEY = 0x0312;
    public const uint MOD_CONTROL = 0x0002;
    public const uint MOD_SHIFT = 0x0004;
    public const uint VK_H = 0x48;

    // Console window hide
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    public const int SW_HIDE = 0;

    // Extract icon from DLL
    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    public static extern uint ExtractIconEx(string lpszFile, int nIconIndex, IntPtr[] phiconLarge, IntPtr[] phiconSmall, uint nIcons);
    [DllImport("user32.dll")]
    public static extern bool DestroyIcon(IntPtr hIcon);
}

public class HotkeyForm : Form {
    private const int HOTKEY_ID = 1;
    public event EventHandler HotkeyPressed;

    public HotkeyForm() {
        this.ShowInTaskbar = false;
        this.FormBorderStyle = FormBorderStyle.None;
        this.Size = new System.Drawing.Size(1, 1);
        this.Opacity = 0;
    }

    public void RegisterGlobalHotkey() {
        NativeMethods.RegisterHotKey(this.Handle, HOTKEY_ID,
            NativeMethods.MOD_CONTROL | NativeMethods.MOD_SHIFT, NativeMethods.VK_H);
    }

    public void UnregisterGlobalHotkey() {
        NativeMethods.UnregisterHotKey(this.Handle, HOTKEY_ID);
    }

    protected override void WndProc(ref Message m) {
        if (m.Msg == NativeMethods.WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID) {
            if (HotkeyPressed != null) HotkeyPressed(this, EventArgs.Empty);
        }
        base.WndProc(ref m);
    }
}
"@


# ============================================================
#  GLOBALS
# ============================================================
$script:ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$script:SettingsFile = Join-Path $script:ScriptDir "clipboard_settings.ini"
$script:HistoryFile = Join-Path $script:ScriptDir "clipboard_data.txt"
$script:ImageFolder = Join-Path $script:ScriptDir "clipboard_images"
$script:ClipHistory = [System.Collections.ArrayList]::new()
$script:MaxHistory = 0
$script:CloseOnCopy = $true
$global:LastClipText = ""
$global:LastImageFingerprint = ""
$script:IsVisible = $false
$script:GuiW = 520
$script:GuiH = 630
$global:SkipClipCheck = $false

# GUI control references
$script:Form = $null
$script:ListPanel = $null
$script:ListContainer = $null
$script:ScrollTrack = $null
$script:ScrollThumb = $null
$script:ThumbDragging = $false
$script:SearchBox = $null
$script:TypeFilter = "all"
$script:ThumbCache = @{}
$script:CachedDisplayList = $null
$script:DisplayListDirty = $true
$script:StatusLabel = $null
$script:TxtStatsLabel = $null
$script:ImgStatsLabel = $null
$script:AllStatsLabel = $null

# ============================================================
#  SETTINGS
# ============================================================
function Save-Settings {
    $cc = "0"
    if ($script:CloseOnCopy) { $cc = "1" }
    $content = "historyFile=" + $script:HistoryFile + "`r`n" +
               "imageFolder=" + $script:ImageFolder + "`r`n" +
               "closeOnCopy=" + $cc + "`r`n" +
               "maxHistory=" + $script:MaxHistory
    [System.IO.File]::WriteAllText($script:SettingsFile, $content, [System.Text.Encoding]::UTF8)
}

function Get-DefaultScreenshotsFolder {
    # Try OneDrive Screenshots first
    if ($env:OneDrive) {
        $oneDriveSS = Join-Path $env:OneDrive "Pictures\Screenshots"
        if (Test-Path (Split-Path $oneDriveSS)) {
            if (-not (Test-Path $oneDriveSS)) { New-Item -ItemType Directory -Path $oneDriveSS -Force | Out-Null }
            return $oneDriveSS
        }
    }
    # Try standard Pictures\Screenshots
    $myPics = [System.Environment]::GetFolderPath("MyPictures")
    if ($myPics -and (Test-Path $myPics)) {
        $stdSS = Join-Path $myPics "Screenshots"
        if (-not (Test-Path $stdSS)) { New-Item -ItemType Directory -Path $stdSS -Force | Out-Null }
        return $stdSS
    }
    return $null
}



function Load-Settings {
    $script:HistoryFile = Join-Path $script:ScriptDir "clipboard_data.txt"
    $script:ImageFolder = Join-Path $script:ScriptDir "clipboard_images"
    $script:CloseOnCopy = $false
    $script:MaxHistory = 0
    $firstRun = -not (Test-Path $script:SettingsFile)

    if (-not $firstRun) {
        try {
            $lines = Get-Content $script:SettingsFile -Encoding UTF8
            foreach ($line in $lines) {
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                $eq = $line.IndexOf("=")
                if ($eq -lt 1) { continue }
                $key = $line.Substring(0, $eq)
                $val = $line.Substring($eq + 1)
                switch ($key) {
                    "historyFile"  { if ($val) { $script:HistoryFile = $val } }
                    "imageFolder"  { if ($val) { $script:ImageFolder = $val } }
                    "closeOnCopy"  { $script:CloseOnCopy = ($val -eq "1") }
                    "maxHistory"   { $script:MaxHistory = [int]$val }
                }
            }
        }
        catch { }
        # If imageFolder wasn't in settings, derive from historyFile location
        $defaultImgFolder = Join-Path $script:ScriptDir "clipboard_images"
        if ($script:ImageFolder -eq $defaultImgFolder) {
            $histDir = Split-Path -Parent $script:HistoryFile
            if ($histDir -ne $script:ScriptDir) {
                $script:ImageFolder = $histDir
                Save-Settings
            }
        }
    }
    else {
        # First run: detect Windows Screenshots folder
        $ssFolder = Get-DefaultScreenshotsFolder
        if ($ssFolder) {
            $script:ImageFolder = $ssFolder
            $script:HistoryFile = Join-Path $ssFolder "clipboard_data.txt"
        }
        if (-not (Test-Path $script:ImageFolder)) {
            New-Item -ItemType Directory -Path $script:ImageFolder -Force | Out-Null
        }
        Save-Settings
    }
}

# ============================================================
#  HISTORY PERSISTENCE
# ============================================================
function Save-History {
    $script:DisplayListDirty = $true
    try {
        if (Test-Path $script:HistoryFile) { Remove-Item $script:HistoryFile -Force }
        if ($script:ClipHistory.Count -eq 0) { return }

        $sb = New-Object System.Text.StringBuilder
        foreach ($entry in $script:ClipHistory) {
            if ($entry.Type -eq "image") {
                [void]$sb.AppendLine($entry.Time + "|||<<IMG>>" + $entry.Image + "|||" + $entry.Dims)
            }
            else {
                $safe = $entry.Text -replace "`r`n","<<NL>>" -replace "`n","<<NL>>"
                [void]$sb.AppendLine($entry.Time + "|||" + $safe)
            }
        }
        $enc = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($script:HistoryFile, $sb.ToString(), $enc)
    }
    catch { }
}

function Load-History {
    $script:ClipHistory.Clear()
    if (-not (Test-Path $script:HistoryFile)) { return }
    try {
        $lines = [System.IO.File]::ReadAllLines($script:HistoryFile, [System.Text.Encoding]::UTF8)
        foreach ($line in $lines) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            $sep = $line.IndexOf("|||")
            if ($sep -lt 1) { continue }
            $timeStr = $line.Substring(0, $sep)
            $rest = $line.Substring($sep + 3)

            if ($rest.StartsWith("<<IMG>>")) {
                $imgRest = $rest.Substring(7)
                $sep2 = $imgRest.IndexOf("|||")
                if ($sep2 -gt 0) {
                    $imgPath = $imgRest.Substring(0, $sep2)
                    $dims = $imgRest.Substring($sep2 + 3)
                }
                else {
                    $imgPath = $imgRest
                    $dims = ""
                }
                if (Test-Path $imgPath) {
                    $e = @{ Type="image"; Time=$timeStr; Text="[IMG] Screenshot $dims"; Image=$imgPath; Dims=$dims }
                    [void]$script:ClipHistory.Add($e)
                }
            }
            else {
                $text = $rest -replace "<<NL>>","`r`n"
                $e = @{ Type="text"; Time=$timeStr; Text=$text; Image=""; Dims="" }
                [void]$script:ClipHistory.Add($e)
            }
        }
    }
    catch { }
}

# ============================================================
#  HELPERS
# ============================================================
function Format-FileSize([long]$bytes) {
    if ($bytes -lt 1024) { return "$bytes B" }
    if ($bytes -lt 1048576) { return "$([math]::Round($bytes / 1024, 1)) KB" }
    return "$([math]::Round($bytes / 1048576, 1)) MB"
}

function Update-StorageStats {
    $imgSize = [long]0; $imgCount = 0; $txtCount = 0
    $txtSize = [long]0
    if (Test-Path $script:ImageFolder) {
        $files = Get-ChildItem $script:ImageFolder -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -match '\.(png|jpg|jpeg|bmp)$' }
        foreach ($fi in $files) { $imgSize += $fi.Length; $imgCount++ }
    }
    if (Test-Path $script:HistoryFile) { $txtSize = (Get-Item $script:HistoryFile).Length }
    $txtCount = @($script:ClipHistory | Where-Object { $_.Type -eq "text" }).Count
    $imgHistCount = @($script:ClipHistory | Where-Object { $_.Type -eq "image" }).Count
    $totalSize = $imgSize + $txtSize
    $totalCount = $imgCount + $txtCount
    if ($null -ne $script:TxtStatsLabel)  { $script:TxtStatsLabel.Text  = "$txtCount entries" }
    if ($null -ne $script:ImgStatsLabel)  { $script:ImgStatsLabel.Text  = "$imgCount images" }
    if ($null -ne $script:AllStatsLabel)  { $script:AllStatsLabel.Text  = "$totalCount files" }
    if ($null -ne $script:StatusLabel) { $script:StatusLabel.Text = "$txtCount entries  |  $imgCount screenshots" }
}

function Get-StartupLnkPath {
    $startup = [System.Environment]::GetFolderPath("Startup")
    return Join-Path $startup "ClipboardManager.lnk"
}

function Get-ImageFingerprint([System.Drawing.Image]$img) {
    # Sample 5 pixels to create a content-based fingerprint
    $w = $img.Width; $h = $img.Height
    $bmp = [System.Drawing.Bitmap]$img
    $p1 = $bmp.GetPixel(0, 0).ToArgb()
    $p2 = $bmp.GetPixel([Math]::Min($w - 1, $w / 2), [Math]::Min($h - 1, $h / 2)).ToArgb()
    $p3 = $bmp.GetPixel($w - 1, 0).ToArgb()
    $p4 = $bmp.GetPixel(0, $h - 1).ToArgb()
    $p5 = $bmp.GetPixel($w - 1, $h - 1).ToArgb()
    return "${w}x${h}_${p1}_${p2}_${p3}_${p4}_${p5}"
}

function Copy-ClipEntry {
    param($e, $panel)
    $global:SkipClipCheck = $true
    if ($e.Type -eq "image" -and $e.Image -and (Test-Path $e.Image)) {
        try {
            $bmp = [System.Drawing.Image]::FromFile($e.Image)
            [System.Windows.Forms.Clipboard]::SetImage($bmp)
            $global:LastImageFingerprint = Get-ImageFingerprint $bmp
            $bmp.Dispose()
        } catch { }
    } else {
        [System.Windows.Forms.Clipboard]::SetText($e.Text)
        $global:LastClipText = $e.Text
    }
    $resetTimer = New-Object System.Windows.Forms.Timer
    $resetTimer.Interval = 1200
    $resetTimer.Add_Tick({
        $global:SkipClipCheck = $false
        $this.Stop()
        $this.Dispose()
    })
    $resetTimer.Start()
    if ($script:CloseOnCopy) {
        $script:Form.Hide(); $script:IsVisible = $false
    } elseif ($null -ne $panel) {
        # Green flash on the copied card
        $gColor = [System.Drawing.Color]::FromArgb(0, 220, 110)
        $origBg = $panel.BackColor
        $bdr = 2
        $pw = $panel.Width; $ph = $panel.Height
        $bt = New-Object System.Windows.Forms.Panel; $bt.BackColor = $gColor; $bt.SetBounds(0, 0, $pw, $bdr)
        $bb = New-Object System.Windows.Forms.Panel; $bb.BackColor = $gColor; $bb.SetBounds(0, ($ph - $bdr), $pw, $bdr)
        $bl = New-Object System.Windows.Forms.Panel; $bl.BackColor = $gColor; $bl.SetBounds(0, 0, $bdr, $ph)
        $br = New-Object System.Windows.Forms.Panel; $br.BackColor = $gColor; $br.SetBounds(($pw - $bdr), 0, $bdr, $ph)
        $cl = New-Object System.Windows.Forms.Label
        $cl.Text = "Copied!"
        $cl.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $cl.ForeColor = [System.Drawing.Color]::White
        $cl.BackColor = $gColor
        $cl.TextAlign = "MiddleCenter"
        $cl.Size = New-Object System.Drawing.Size(70, 20)
        $cl.Location = New-Object System.Drawing.Point((($pw - 70) / 2), 4)
        $parts = @($bt, $bb, $bl, $br, $cl)
        foreach ($p in $parts) { $panel.Controls.Add($p); $p.BringToFront() }
        $flashTimer = New-Object System.Windows.Forms.Timer
        $flashTimer.Interval = 400
        $flashTimer.Add_Tick({
            foreach ($p in $parts) {
                $panel.Controls.Remove($p)
                try { $p.Dispose() } catch { }
            }
            $this.Stop(); $this.Dispose()
        }.GetNewClosure())
        $flashTimer.Start()
    }
}

function Apply-TypeFilter {
    if ($null -eq $script:ListPanel) { return }
    $script:ListPanel.SuspendLayout()
    foreach ($ctrl in $script:ListPanel.Controls) {
        if ($script:TypeFilter -eq "all") {
            $ctrl.Visible = $true
        } elseif ($script:TypeFilter -eq "text") {
            $ctrl.Visible = ($ctrl.Name -eq "text")
        } elseif ($script:TypeFilter -eq "image") {
            $ctrl.Visible = ($ctrl.Name -eq "image")
        }
    }
    $script:ListPanel.ResumeLayout()
}

function Refresh-List {
    if ($null -eq $script:ListPanel) { return }
    $script:ListPanel.SuspendLayout()
    foreach ($c in $script:ListPanel.Controls) { try { $c.Dispose() } catch { } }
    $script:ListPanel.Controls.Clear()

    $query = ""
    if ($null -ne $script:SearchBox) { $query = $script:SearchBox.Text }
    $count = 0
    $imgCount = 0
    $sbW = 6
    $visibleW = if ($null -ne $script:ListContainer) { $script:ListContainer.Width - $sbW } else { 494 }
    $panelW = $visibleW - 14
    $rowBg1  = [System.Drawing.Color]::FromArgb(32, 32, 56)
    $rowBg2  = [System.Drawing.Color]::FromArgb(40, 40, 68)
    $hoverBg = [System.Drawing.Color]::FromArgb(55, 55, 90)
    $timeFg  = [System.Drawing.Color]::FromArgb(130, 160, 220)
    $textFg  = [System.Drawing.Color]::FromArgb(210, 225, 210)
    $imgFg   = [System.Drawing.Color]::FromArgb(100, 195, 255)
    $even    = $true

    # Build display list: use cache if available
    if ($script:DisplayListDirty -or $null -eq $script:CachedDisplayList) {
        $trackedPaths = @{}
        foreach ($e in $script:ClipHistory) {
            if ($e.Type -eq "image" -and $e.Image) { $trackedPaths[$e.Image] = $true }
        }
        $displayList = [System.Collections.ArrayList]::new($script:ClipHistory)
        if (Test-Path $script:ImageFolder) {
            $folderImgs = Get-ChildItem $script:ImageFolder -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -match '\.(png|jpg|jpeg|bmp)$' } |
                Sort-Object CreationTime -Descending
            foreach ($fi in $folderImgs) {
                if ($trackedPaths.ContainsKey($fi.FullName)) { continue }
                $ts = $fi.CreationTime.ToString("MM/dd/yyyy HH:mm")
                $folderEntry = @{ Type="image"; Time=$ts; Text="[IMG] $($fi.Name)"; Image=$fi.FullName; Dims=""; FolderOnly=$true }
                [void]$displayList.Add($folderEntry)
            }
        }
        $script:CachedDisplayList = $displayList
        $script:DisplayListDirty = $false
    } else {
        $displayList = $script:CachedDisplayList
    }

    for ($idx = 0; $idx -lt $displayList.Count; $idx++) {
        $entry = $displayList[$idx]
        if ($entry.Type -eq "image") { $imgCount++ }

        $searchText = $entry.Text
        if ($query -and ($searchText -notlike "*$query*") -and ($entry.Time -notlike "*$query*")) { continue }

        $bg = if ($even) { $rowBg1 } else { $rowBg2 }
        $even = -not $even

        $ep = New-Object System.Windows.Forms.Panel
        $ep.Width = $panelW
        $ep.BackColor = $bg
        $ep.Cursor = [System.Windows.Forms.Cursors]::Hand
        $ep.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 2)
        $ep.Tag = $idx
        $ep.Name = $entry.Type

        # Left color border: green=text, blue=image
        $border = New-Object System.Windows.Forms.Panel
        $border.Size = New-Object System.Drawing.Size(4, 300)
        $border.Location = New-Object System.Drawing.Point(0, 0)
        if ($entry.Type -eq "image") { $border.BackColor = [System.Drawing.Color]::FromArgb(60, 160, 255) }
        else { $border.BackColor = [System.Drawing.Color]::FromArgb(0, 210, 110) }
        $ep.Controls.Add($border)

        # Time and date labels - split into corners
        $timeParts = $entry.Time -split ' '
        $datePart = if ($timeParts.Count -ge 1) { $timeParts[0] } else { "" }
        $timePart = if ($timeParts.Count -ge 2) { $timeParts[1] } else { "" }

        $tl = New-Object System.Windows.Forms.Label
        $tl.Text = $timePart
        $tl.Location = New-Object System.Drawing.Point(12, 5)
        $tl.Size = New-Object System.Drawing.Size(80, 20)
        $tl.ForeColor = $timeFg
        $tl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $tl.BackColor = [System.Drawing.Color]::Transparent
        $tl.TextAlign = "MiddleLeft"
        $ep.Controls.Add($tl)

        $dl = New-Object System.Windows.Forms.Label
        $dl.Text = $datePart
        $dl.Location = New-Object System.Drawing.Point(($panelW - 110), 5)
        $dl.Size = New-Object System.Drawing.Size(100, 20)
        $dl.ForeColor = $timeFg
        $dl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $dl.BackColor = [System.Drawing.Color]::Transparent
        $dl.TextAlign = "MiddleRight"
        $ep.Controls.Add($dl)

        if ($entry.Type -eq "image" -and $entry.Image -and (Test-Path $entry.Image)) {
            $thumbW = [int]($panelW * 0.72)
            try {
                $imgPath = $entry.Image
                if ($script:ThumbCache.ContainsKey($imgPath)) {
                    $cachedInfo = $script:ThumbCache[$imgPath]
                    $thumb = $cachedInfo.Thumb
                    $thumbW = $cachedInfo.W
                    $thumbH = $cachedInfo.H
                } else {
                    $img = [System.Drawing.Image]::FromFile($imgPath)
                    $ratio = $thumbW / $img.Width
                    $thumbH = [int]($img.Height * $ratio)
                    if ($thumbH -gt 200) { $thumbH = 200; $ratio = $thumbH / $img.Height; $thumbW = [int]($img.Width * $ratio) }
                    $thumb = $img.GetThumbnailImage($thumbW, $thumbH, $null, [IntPtr]::Zero)
                    $img.Dispose()
                    $script:ThumbCache[$imgPath] = @{ Thumb=$thumb; W=$thumbW; H=$thumbH }
                }

                $pb = New-Object System.Windows.Forms.PictureBox
                $pb.Image = $thumb
                $pb.SizeMode = "Zoom"
                $pb.Size = New-Object System.Drawing.Size($thumbW, $thumbH)
                $pb.Location = New-Object System.Drawing.Point(12, 28)
                $pb.BackColor = [System.Drawing.Color]::Black
                $pb.Cursor = [System.Windows.Forms.Cursors]::Hand
                $ep.Controls.Add($pb)

                $dl = New-Object System.Windows.Forms.Label
                $dl.Text = $entry.Dims
                $dl.Location = New-Object System.Drawing.Point(($thumbW + 18), 30)
                $dl.AutoSize = $true
                $dl.ForeColor = $imgFg
                $dl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
                $dl.BackColor = [System.Drawing.Color]::Transparent
                $ep.Controls.Add($dl)

                $ep.Height = $thumbH + 36
                $border.Height = $ep.Height
            }
            catch {
                $fl = New-Object System.Windows.Forms.Label
                $fl.Text = "Screenshot " + $entry.Dims
                $fl.Location = New-Object System.Drawing.Point(12, 28)
                $fl.AutoSize = $true
                $fl.ForeColor = $imgFg
                $fl.Font = New-Object System.Drawing.Font("Segoe UI", 10)
                $fl.BackColor = [System.Drawing.Color]::Transparent
                $ep.Controls.Add($fl)
                $ep.Height = 52
                $border.Height = 52
            }
        }
        else {
            $flatText = $entry.Text -replace "`r`n"," | " -replace "`n"," | " -replace "`t"," "
            $shortText = $flatText
            if ($shortText.Length -gt 140) { $shortText = $shortText.Substring(0, 137) + "..." }
            $origText = $entry.Text -replace "`t","    "
            $cl = New-Object System.Windows.Forms.Label
            $cl.Text = $shortText
            $cl.Location = New-Object System.Drawing.Point(12, 28)
            $cl.Size = New-Object System.Drawing.Size(($panelW - 24), 36)
            $cl.ForeColor = $textFg
            $cl.Font = New-Object System.Drawing.Font("Segoe UI", 10)
            $cl.BackColor = [System.Drawing.Color]::Transparent
            $isMultiLine = $entry.Text.Contains([char]10)
            $cl.Tag = @{ Full=$origText; Short=$shortText; Expanded=$false; CollapsedH=68; IsLong=($flatText.Length -gt 140 -or $isMultiLine) }
            $ep.Controls.Add($cl)
            $ep.Height = 68
            $border.Height = 68

            # Single-click expand/collapse for long text
            if ($cl.Tag.IsLong) {
                $expandHandler = {
                    $lbl = $this
                    $info = $lbl.Tag
                    $parentPanel = $lbl.Parent
                    $bdr = $null
                    foreach ($c in $parentPanel.Controls) {
                        if ($c -is [System.Windows.Forms.Panel] -and $c.Width -eq 4) { $bdr = $c; break }
                    }
                    if (-not $info.Expanded) {
                        $lbl.Text = $info.Full
                        $g = [System.Windows.Forms.TextRenderer]::MeasureText($info.Full, $lbl.Font, (New-Object System.Drawing.Size($lbl.Width, 0)), "WordBreak")
                        $newLblH = [Math]::Max(36, $g.Height + 4)
                        $lbl.Height = $newLblH
                        $newH = $newLblH + 36
                        $parentPanel.Height = $newH
                        if ($bdr) { $bdr.Height = $newH }
                        $info.Expanded = $true
                    } else {
                        $lbl.Text = $info.Short
                        $lbl.Height = 36
                        $parentPanel.Height = $info.CollapsedH
                        if ($bdr) { $bdr.Height = $info.CollapsedH }
                        $info.Expanded = $false
                    }
                }.GetNewClosure()
                $cl.Add_Click($expandHandler)
            }
        }

        # Hover
        $hOn = { $this.Parent.BackColor = $hoverBg }.GetNewClosure()
        $hOff = { $this.Parent.BackColor = $bg }.GetNewClosure()
        foreach ($child in $ep.Controls) {
            $child.Add_MouseEnter($hOn)
            $child.Add_MouseLeave($hOff)
        }
        $ep.Add_MouseEnter({ $this.BackColor = $hoverBg }.GetNewClosure())
        $ep.Add_MouseLeave({ $this.BackColor = $bg }.GetNewClosure())

        # Double-click to copy
        $capturedCopyEntry = $entry
        $capturedPanel = $ep
        $copyAction = {
            Copy-ClipEntry $capturedCopyEntry $capturedPanel
        }.GetNewClosure()
        $ep.Add_DoubleClick($copyAction)
        foreach ($child in $ep.Controls) { $child.Add_DoubleClick($copyAction) }

        # Right-click context menu
        $ctx = New-Object System.Windows.Forms.ContextMenuStrip
        $ctx.BackColor = [System.Drawing.Color]::FromArgb(35, 35, 60)
        $ctx.ForeColor = [System.Drawing.Color]::White
        $ctx.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $ctx.ShowImageMargin = $false

        if ($entry.Type -eq "image" -and $entry.Image -and (Test-Path $entry.Image)) {
            $imgPath = $entry.Image
            $openFull = $ctx.Items.Add("Open Full Size")
            $openFull.Add_Click({
                try { [System.Diagnostics.Process]::Start($imgPath) | Out-Null } catch { }
            }.GetNewClosure())

            $openLoc = $ctx.Items.Add("Open File Location")
            $openLoc.Add_Click({
                try { [System.Diagnostics.Process]::Start("explorer.exe", "/select,`"$imgPath`"") | Out-Null } catch { }
            }.GetNewClosure())
            [void]$ctx.Items.Add("-")
        }

        $delItem = $ctx.Items.Add("Delete")
        $delItem.ForeColor = [System.Drawing.Color]::FromArgb(255, 100, 100)
        $capturedEntry = $entry
        $capturedPanel = $ep
        $capturedStatusLabel = $script:StatusLabel
        $capturedHistory = $script:ClipHistory
        $isFolderOnly = $entry.FolderOnly -eq $true
        $delItem.Add_Click({
            $e = $capturedEntry
            if ($e.Type -eq "image" -and $e.Image -and (Test-Path $e.Image)) {
                Remove-Item $e.Image -Force -ErrorAction SilentlyContinue
                $script:ThumbCache.Remove($e.Image)
            }
            $script:DisplayListDirty = $true
            if (-not $isFolderOnly) {
                $capturedHistory.Remove($e)
                Save-History
            }
            # Remove just this panel (no flicker)
            $script:ListPanel.SuspendLayout()
            $script:ListPanel.Controls.Remove($capturedPanel)
            try { $capturedPanel.Dispose() } catch { }
            $script:ListPanel.ResumeLayout()
            # Update all stats and counter
            Update-StorageStats
        }.GetNewClosure())

        $ep.ContextMenuStrip = $ctx
        foreach ($child in $ep.Controls) { $child.ContextMenuStrip = $ctx }

        [void]$script:ListPanel.Controls.Add($ep)
        $count++
    }

    $script:ListPanel.ResumeLayout()
    if ($null -ne $script:StatusLabel) { $script:StatusLabel.Text = "$count items  |  $imgCount screenshots" }
    Update-StorageStats
    Apply-TypeFilter
}

function Enforce-MaxHistory {
    if ($script:MaxHistory -gt 0) {
        while ($script:ClipHistory.Count -gt $script:MaxHistory) {
            $old = $script:ClipHistory[$script:ClipHistory.Count - 1]
            if ($old.Type -eq "image" -and $old.Image -and (Test-Path $old.Image)) {
                Remove-Item $old.Image -Force -ErrorAction SilentlyContinue
            }
            $script:ClipHistory.RemoveAt($script:ClipHistory.Count - 1)
        }
    }
}

function Toggle-Gui {
    if ($script:IsVisible) {
        $script:Form.Hide()
        $script:IsVisible = $false
    }
    else {
        $script:DisplayListDirty = $true
        Refresh-List
        $scrW = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width
        $script:Form.Location = New-Object System.Drawing.Point((($scrW - $script:GuiW) / 2), 10)
        $script:Form.Show()
        $script:Form.Activate()
        if ($null -ne $script:SearchBox) { $script:SearchBox.Focus() }
        $script:IsVisible = $true
    }
}

# ============================================================
#  BUILD GUI
# ============================================================
function Build-MainForm {
    Load-Settings
    if (-not (Test-Path $script:ImageFolder)) {
        New-Item -ItemType Directory -Path $script:ImageFolder -Force | Out-Null
    }
    Load-History

    # Snapshot current clipboard to prevent re-adding on startup
    try {
        if ([System.Windows.Forms.Clipboard]::ContainsImage()) {
            $cbImg = [System.Windows.Forms.Clipboard]::GetImage()
            if ($null -ne $cbImg) {
                $global:LastImageFingerprint = Get-ImageFingerprint $cbImg
                $cbImg.Dispose()
            }
        }
        if ([System.Windows.Forms.Clipboard]::ContainsText()) {
            $global:LastClipText = [System.Windows.Forms.Clipboard]::GetText()
        }
    } catch { }

    $darkBg  = [System.Drawing.Color]::FromArgb(26, 26, 46)
    $darkBg2 = [System.Drawing.Color]::FromArgb(42, 42, 74)
    $titleBg = [System.Drawing.Color]::FromArgb(22, 33, 62)
    $green   = [System.Drawing.Color]::FromArgb(0, 255, 136)
    $gray    = [System.Drawing.Color]::FromArgb(187, 187, 187)
    $white   = [System.Drawing.Color]::White

    # --- Main Form ---
    $script:Form = New-Object System.Windows.Forms.Form
    $script:Form.Text = "Clipboard Manager"
    $script:Form.Size = New-Object System.Drawing.Size($script:GuiW, $script:GuiH)
    $script:Form.FormBorderStyle = "None"
    $script:Form.BackColor = $darkBg
    $script:Form.TopMost = $true
    $script:Form.ShowInTaskbar = $false
    $script:Form.StartPosition = "Manual"
    $scrW = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width
    $script:Form.Location = New-Object System.Drawing.Point((($scrW - $script:GuiW) / 2), 10)
    $f = $script:Form

    $script:Form.Add_Shown({
        $style = [NativeMethods]::GetWindowLong($this.Handle, [NativeMethods]::GWL_EXSTYLE)
        [NativeMethods]::SetWindowLong($this.Handle, [NativeMethods]::GWL_EXSTYLE,
            $style -bor [NativeMethods]::WS_EX_TOOLWINDOW) | Out-Null
    })

    # --- Title Bar ---
    $titleBar = New-Object System.Windows.Forms.Panel
    $titleBar.Size = New-Object System.Drawing.Size(520, 35)
    $titleBar.Location = New-Object System.Drawing.Point(0, 0)
    $titleBar.BackColor = $titleBg

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Clipboard Manager"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = $green
    $titleLabel.Size = New-Object System.Drawing.Size(520, 35)
    $titleLabel.TextAlign = "MiddleCenter"
    $titleBar.Controls.Add($titleLabel)

    $dragHandler = {
        [NativeMethods]::ReleaseCapture() | Out-Null
        [NativeMethods]::SendMessage($f.Handle, [NativeMethods]::WM_NCLBUTTONDOWN,
            [NativeMethods]::HT_CAPTION, 0) | Out-Null
    }
    $titleBar.Add_MouseDown($dragHandler)
    $titleLabel.Add_MouseDown($dragHandler)
    $f.Controls.Add($titleBar)

    # --- Close Button ---
    $closeBtn = New-Object System.Windows.Forms.Button
    $closeBtn.Text = "X"
    $closeBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $closeBtn.Size = New-Object System.Drawing.Size(40, 35)
    $closeBtn.Location = New-Object System.Drawing.Point(480, 0)
    $closeBtn.FlatStyle = "Flat"
    $closeBtn.FlatAppearance.BorderSize = 0
    $closeBtn.BackColor = $titleBg
    $closeBtn.ForeColor = [System.Drawing.Color]::FromArgb(255, 100, 100)
    $closeBtn.Add_Click({ $f.Hide(); $script:IsVisible = $false })
    $titleBar.Controls.Add($closeBtn)
    $closeBtn.BringToFront()

    # --- Search ---
    $searchLabel = New-Object System.Windows.Forms.Label
    $searchLabel.Text = "Search:"
    $searchLabel.Location = New-Object System.Drawing.Point(12, 45)
    $searchLabel.Size = New-Object System.Drawing.Size(55, 26)
    $searchLabel.ForeColor = $gray
    $searchLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $searchLabel.TextAlign = "MiddleLeft"
    $f.Controls.Add($searchLabel)

    $script:SearchBox = New-Object System.Windows.Forms.TextBox
    $script:SearchBox.Location = New-Object System.Drawing.Point(68, 47)
    $script:SearchBox.Size = New-Object System.Drawing.Size(244, 26)
    $script:SearchBox.BackColor = $darkBg2
    $script:SearchBox.ForeColor = $green
    $script:SearchBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $script:SearchBox.BorderStyle = "FixedSingle"
    $script:SearchBox.ShortcutsEnabled = $true
    # Debounced search: wait 300ms after last keystroke before refreshing
    $script:SearchTimer = New-Object System.Windows.Forms.Timer
    $script:SearchTimer.Interval = 300
    $script:SearchTimer.Add_Tick({
        $script:SearchTimer.Stop()
        Refresh-List
    })
    $script:SearchBox.Add_TextChanged({
        $script:SearchTimer.Stop()
        $script:SearchTimer.Start()
    })
    $script:SearchBox.Add_KeyDown({
        param($sender, $e)
        if ($e.Control -and $e.KeyCode -eq "A") {
            $sender.SelectAll()
            $e.SuppressKeyPress = $true
        }
        elseif ($e.Control -and $e.KeyCode -eq "Back") {
            $e.SuppressKeyPress = $true
            $pos = $sender.SelectionStart
            if ($pos -gt 0) {
                $txt = $sender.Text
                $i = $pos - 1
                while ($i -gt 0 -and $txt[$i - 1] -eq ' ') { $i-- }
                while ($i -gt 0 -and $txt[$i - 1] -ne ' ') { $i-- }
                $sender.Text = $txt.Substring(0, $i) + $txt.Substring($pos)
                $sender.SelectionStart = $i
            }
        }
    })
    $f.Controls.Add($script:SearchBox)

    # --- Search clear X button ---
    $searchClearBtn = New-Object System.Windows.Forms.Button
    $searchClearBtn.Text = [char]0x2715
    $searchClearBtn.Location = New-Object System.Drawing.Point(312, 47)
    $searchClearBtn.Size = New-Object System.Drawing.Size(26, 26)
    $searchClearBtn.FlatStyle = "Flat"
    $searchClearBtn.FlatAppearance.BorderSize = 0
    $searchClearBtn.BackColor = $darkBg2
    $searchClearBtn.ForeColor = [System.Drawing.Color]::FromArgb(180, 180, 180)
    $searchClearBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $searchClearBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $searchClearBtn.Add_Click({
        $script:SearchBox.Text = ""
        $script:SearchBox.Focus()
    })
    $f.Controls.Add($searchClearBtn)

    # --- Filter buttons ---
    $filterActiveBg = [System.Drawing.Color]::FromArgb(0, 170, 90)
    $filterInactiveBg = $darkBg
    $filterBorder = [System.Drawing.Color]::FromArgb(80, 80, 120)

    $script:TextFilterBtn = New-Object System.Windows.Forms.Button
    $script:TextFilterBtn.Text = "Text only"
    $script:TextFilterBtn.Location = New-Object System.Drawing.Point(345, 47)
    $script:TextFilterBtn.Size = New-Object System.Drawing.Size(78, 26)
    $script:TextFilterBtn.FlatStyle = "Flat"
    $script:TextFilterBtn.FlatAppearance.BorderColor = $filterBorder
    $script:TextFilterBtn.BackColor = $filterInactiveBg
    $script:TextFilterBtn.ForeColor = $gray
    $script:TextFilterBtn.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $script:TextFilterBtn.Add_Click({
        if ($script:TypeFilter -eq "text") {
            $script:TypeFilter = "all"
            $script:TextFilterBtn.BackColor = $filterInactiveBg
            $script:TextFilterBtn.ForeColor = $gray
            $script:TextFilterBtn.FlatAppearance.BorderColor = $filterBorder
        } else {
            $script:TypeFilter = "text"
            $script:TextFilterBtn.BackColor = $filterActiveBg
            $script:TextFilterBtn.ForeColor = $white
            $script:TextFilterBtn.FlatAppearance.BorderColor = $filterActiveBg
            $script:ImgFilterBtn.BackColor = $filterInactiveBg
            $script:ImgFilterBtn.ForeColor = $gray
            $script:ImgFilterBtn.FlatAppearance.BorderColor = $filterBorder
        }
        Apply-TypeFilter
    })
    $f.Controls.Add($script:TextFilterBtn)

    $script:ImgFilterBtn = New-Object System.Windows.Forms.Button
    $script:ImgFilterBtn.Text = "Screenshots"
    $script:ImgFilterBtn.Location = New-Object System.Drawing.Point(427, 47)
    $script:ImgFilterBtn.Size = New-Object System.Drawing.Size(82, 26)
    $script:ImgFilterBtn.FlatStyle = "Flat"
    $script:ImgFilterBtn.FlatAppearance.BorderColor = $filterBorder
    $script:ImgFilterBtn.BackColor = $filterInactiveBg
    $script:ImgFilterBtn.ForeColor = $gray
    $script:ImgFilterBtn.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $script:ImgFilterBtn.Add_Click({
        if ($script:TypeFilter -eq "image") {
            $script:TypeFilter = "all"
            $script:ImgFilterBtn.BackColor = $filterInactiveBg
            $script:ImgFilterBtn.ForeColor = $gray
            $script:ImgFilterBtn.FlatAppearance.BorderColor = $filterBorder
        } else {
            $script:TypeFilter = "image"
            $script:ImgFilterBtn.BackColor = [System.Drawing.Color]::FromArgb(40, 120, 220)
            $script:ImgFilterBtn.ForeColor = $white
            $script:ImgFilterBtn.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(40, 120, 220)
            $script:TextFilterBtn.BackColor = $filterInactiveBg
            $script:TextFilterBtn.ForeColor = $gray
            $script:TextFilterBtn.FlatAppearance.BorderColor = $filterBorder
        }
        Apply-TypeFilter
    })
    $f.Controls.Add($script:ImgFilterBtn)

    # --- Hint ---
    $hint = New-Object System.Windows.Forms.Label
    $hint.Text = "Double-click to copy"
    $hint.Location = New-Object System.Drawing.Point(0, 78)
    $hint.Size = New-Object System.Drawing.Size(520, 20)
    $hint.ForeColor = [System.Drawing.Color]::FromArgb(0, 204, 102)
    $hint.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $hint.TextAlign = "MiddleCenter"
    $f.Controls.Add($hint)

    # --- Scrollable List ---
    $sbW = 6
    $script:ListContainer = New-Object System.Windows.Forms.Panel
    $script:ListContainer.Location = New-Object System.Drawing.Point(10, 100)
    $script:ListContainer.Size = New-Object System.Drawing.Size(500, 380)
    $script:ListContainer.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 48)
    $f.Controls.Add($script:ListContainer)

    $nativeSbW = [System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth
    $script:ListPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $script:ListPanel.Location = New-Object System.Drawing.Point(0, 0)
    $script:ListPanel.Size = New-Object System.Drawing.Size((500 + $nativeSbW), 380)
    $script:ListPanel.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 48)
    $script:ListPanel.AutoScroll = $true
    $script:ListPanel.FlowDirection = "TopDown"
    $script:ListPanel.WrapContents = $false
    $script:ListPanel.BorderStyle = "None"
    $script:ListContainer.Controls.Add($script:ListPanel)

    # Custom scrollbar track
    $script:ScrollTrack = New-Object System.Windows.Forms.Panel
    $script:ScrollTrack.Size = New-Object System.Drawing.Size($sbW, 380)
    $script:ScrollTrack.Location = New-Object System.Drawing.Point((500 - $sbW), 0)
    $script:ScrollTrack.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 52)
    $script:ListContainer.Controls.Add($script:ScrollTrack)
    $script:ScrollTrack.BringToFront()

    # Custom scrollbar thumb
    $script:ScrollThumb = New-Object System.Windows.Forms.Panel
    $script:ScrollThumb.Size = New-Object System.Drawing.Size($sbW, 50)
    $script:ScrollThumb.Location = New-Object System.Drawing.Point(0, 0)
    $script:ScrollThumb.BackColor = [System.Drawing.Color]::FromArgb(0, 200, 110)
    $script:ScrollThumb.Visible = $false
    $script:ScrollTrack.Controls.Add($script:ScrollThumb)

    # Thumb drag support
    $script:ThumbDragging = $false
    $script:ThumbDragStartY = 0
    $script:ThumbDragStartTop = 0

    $script:ScrollThumb.Add_MouseDown({
        param($s, $e)
        if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
        $script:ThumbDragging = $true
        $script:ThumbDragStartY = [System.Windows.Forms.Cursor]::Position.Y
        $script:ThumbDragStartTop = $script:ScrollThumb.Top
    })
    $script:ScrollThumb.Add_MouseMove({
        param($s, $e)
        if (-not $script:ThumbDragging) { return }
        $screenY = [System.Windows.Forms.Cursor]::Position.Y
        $delta = $screenY - $script:ThumbDragStartY
        $trackH = $script:ScrollTrack.Height
        $thumbH = $script:ScrollThumb.Height
        $maxThumbY = $trackH - $thumbH
        $newY = $script:ThumbDragStartTop + $delta
        $newY = [Math]::Max(0, [Math]::Min($newY, $maxThumbY))
        $script:ScrollThumb.Top = $newY
        # Sync FlowLayoutPanel scroll
        $totalH = $script:ListPanel.DisplayRectangle.Height
        $visH = $script:ListPanel.ClientSize.Height
        $maxScroll = $totalH - $visH
        if ($maxThumbY -gt 0 -and $maxScroll -gt 0) {
            $scrollY = [int](($newY / $maxThumbY) * $maxScroll)
            $script:ListPanel.AutoScrollPosition = New-Object System.Drawing.Point(0, $scrollY)
        }
    })
    $script:ScrollThumb.Add_MouseUp({ $script:ThumbDragging = $false })

    # Click track to jump
    $script:ScrollTrack.Add_MouseClick({
        param($s, $e)
        $trackH = $script:ScrollTrack.Height
        $thumbH = $script:ScrollThumb.Height
        $maxThumbY = $trackH - $thumbH
        $newY = [Math]::Max(0, [Math]::Min(($e.Y - $thumbH / 2), $maxThumbY))
        $script:ScrollThumb.Top = $newY
        $totalH = $script:ListPanel.DisplayRectangle.Height
        $visH = $script:ListPanel.ClientSize.Height
        $maxScroll = $totalH - $visH
        if ($maxThumbY -gt 0 -and $maxScroll -gt 0) {
            $scrollY = [int](($newY / $maxThumbY) * $maxScroll)
            $script:ListPanel.AutoScrollPosition = New-Object System.Drawing.Point(0, $scrollY)
        }
    })

    # Scroll sync timer
    $scrollSync = New-Object System.Windows.Forms.Timer
    $scrollSync.Interval = 30
    $scrollSync.Add_Tick({
        if ($null -eq $script:ListPanel -or $null -eq $script:ScrollThumb) { return }
        $totalH = $script:ListPanel.DisplayRectangle.Height
        $visH = $script:ListPanel.ClientSize.Height
        $trackH = $script:ScrollTrack.Height
        if ($totalH -le $visH) {
            $script:ScrollThumb.Visible = $false
            return
        }
        $script:ScrollThumb.Visible = $true
        $thumbH = [Math]::Max(20, [int]($trackH * $visH / $totalH))
        $script:ScrollThumb.Height = $thumbH
        if (-not $script:ThumbDragging) {
            $scrollY = -$script:ListPanel.AutoScrollPosition.Y
            $maxScroll = $totalH - $visH
            if ($maxScroll -gt 0) {
                $thumbY = [int](($scrollY / $maxScroll) * ($trackH - $thumbH))
                $thumbY = [Math]::Max(0, [Math]::Min($thumbY, $trackH - $thumbH))
                $script:ScrollThumb.Top = $thumbY
            }
        }
    })
    $scrollSync.Start()

    # --- PATH label (single line) + Change button ---
    $pathLabel = New-Object System.Windows.Forms.Label
    $pText = $script:ImageFolder
    if ($pText.Length -gt 55) { $pText = "..." + $pText.Substring($pText.Length - 52) }
    $pathLabel.Text = "PATH: " + $pText
    $pathLabel.Location = New-Object System.Drawing.Point(12, 487)
    $pathLabel.Size = New-Object System.Drawing.Size(410, 16)
    $pathLabel.ForeColor = [System.Drawing.Color]::FromArgb(140, 140, 140)
    $pathLabel.Font = New-Object System.Drawing.Font("Segoe UI", 7)
    $pathLabel.TextAlign = "MiddleLeft"
    $f.Controls.Add($pathLabel)

    $changePathBtn = New-Object System.Windows.Forms.Button
    $changePathBtn.Text = "Change..."
    $changePathBtn.Location = New-Object System.Drawing.Point(430, 485)
    $changePathBtn.Size = New-Object System.Drawing.Size(80, 22)
    $changePathBtn.FlatStyle = "Flat"
    $changePathBtn.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(0, 170, 255)
    $changePathBtn.BackColor = $darkBg
    $changePathBtn.ForeColor = [System.Drawing.Color]::FromArgb(0, 170, 255)
    $changePathBtn.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $changePathBtn.Add_Click({
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description = "Choose folder for clipboard history"
        if ($fbd.ShowDialog() -eq "OK") {
            $np = Join-Path $fbd.SelectedPath "clipboard_data.txt"
            $script:ImageFolder = $fbd.SelectedPath
            if (-not (Test-Path $script:ImageFolder)) { New-Item -ItemType Directory -Path $script:ImageFolder -Force | Out-Null }
            if ((Test-Path $script:HistoryFile) -and ($script:HistoryFile -ne $np)) {
                try { Copy-Item $script:HistoryFile $np -Force } catch { }
            }
            $script:HistoryFile = $np
            $t = $script:ImageFolder; if ($t.Length -gt 55) { $t = "..." + $t.Substring($t.Length - 52) }
            $pathLabel.Text = "PATH: " + $t
            Save-Settings; Save-History
        }
    })
    $f.Controls.Add($changePathBtn)

    # --- Checkboxes (same row) ---
    $closeChk = New-Object System.Windows.Forms.CheckBox
    $closeChk.Text = "Close after copy"
    $closeChk.Location = New-Object System.Drawing.Point(15, 510)
    $closeChk.Size = New-Object System.Drawing.Size(155, 22)
    $closeChk.ForeColor = $gray; $closeChk.BackColor = $darkBg
    $closeChk.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $closeChk.Checked = $script:CloseOnCopy
    $closeChk.Add_Click({ $script:CloseOnCopy = $closeChk.Checked; Save-Settings })
    $f.Controls.Add($closeChk)

    $startupChk = New-Object System.Windows.Forms.CheckBox
    $startupChk.Text = "Run at startup"
    $startupChk.Location = New-Object System.Drawing.Point(175, 510)
    $startupChk.Size = New-Object System.Drawing.Size(130, 22)
    $startupChk.ForeColor = $gray; $startupChk.BackColor = $darkBg
    $startupChk.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $startupChk.Checked = (Test-Path (Get-StartupLnkPath))
    $startupChk.Add_Click({
        $lnk = Get-StartupLnkPath
        if ($startupChk.Checked) {
            $ws = New-Object -ComObject WScript.Shell
            $sc = $ws.CreateShortcut($lnk)
            $sc.TargetPath = "powershell.exe"
            $sc.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"" + $PSCommandPath + "`" -Hidden"
            $sc.WorkingDirectory = $script:ScriptDir
            $sc.WindowStyle = 7
            $sc.IconLocation = "shell32.dll,260"
            $sc.Save()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
        }
        else {
            if (Test-Path $lnk) { Remove-Item $lnk -Force }
        }
    })
    $f.Controls.Add($startupChk)

    # Max history (inline with checkboxes)
    $mhLabel = New-Object System.Windows.Forms.Label
    $mhLabel.Text = "Max history:"
    $mhLabel.Location = New-Object System.Drawing.Point(315, 510)
    $mhLabel.Size = New-Object System.Drawing.Size(85, 22)
    $mhLabel.ForeColor = $gray
    $mhLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $mhLabel.TextAlign = "MiddleLeft"
    $f.Controls.Add($mhLabel)

    $mhSpinner = New-Object System.Windows.Forms.NumericUpDown
    $mhSpinner.Location = New-Object System.Drawing.Point(400, 510)
    $mhSpinner.Size = New-Object System.Drawing.Size(75, 24)
    $mhSpinner.Minimum = 0; $mhSpinner.Maximum = 99999
    $mhSpinner.Value = [Math]::Max(0, $script:MaxHistory)
    $mhSpinner.BackColor = $darkBg2; $mhSpinner.ForeColor = $green
    $mhSpinner.Font = New-Object System.Drawing.Font("Consolas", 9)
    $mhSpinner.TextAlign = "Center"
    $mhSpinner.BorderStyle = "FixedSingle"
    $f.Controls.Add($mhSpinner)

    $mhSpinner.Add_ValueChanged({
        $script:MaxHistory = [int]$mhSpinner.Value
        Save-Settings
    })

    $mhHint = New-Object System.Windows.Forms.Label
    $mhHint.Text = "0 = infinity"
    $mhHint.Location = New-Object System.Drawing.Point(400, 534)
    $mhHint.Size = New-Object System.Drawing.Size(75, 14)
    $mhHint.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 120)
    $mhHint.Font = New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Italic)
    $mhHint.TextAlign = "MiddleCenter"
    $f.Controls.Add($mhHint)

    # --- Status row ---
    $script:StatusLabel = New-Object System.Windows.Forms.Label
    $script:StatusLabel.Text = "0 items"
    $script:StatusLabel.Location = New-Object System.Drawing.Point(15, 540)
    $script:StatusLabel.Size = New-Object System.Drawing.Size(300, 20)
    $script:StatusLabel.ForeColor = $white
    $script:StatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $script:StatusLabel.TextAlign = "MiddleLeft"
    $f.Controls.Add($script:StatusLabel)

    # --- Buttons row (horizontal) + stats below each ---
    $btnY = 566; $btnH = 26; $btnW = 110; $gap = 6
    $statsY = $btnY + $btnH + 2

    $cleanTxtBtn = New-Object System.Windows.Forms.Button
    $cleanTxtBtn.Text = "Clean text"
    $cleanTxtBtn.Location = New-Object System.Drawing.Point(15, $btnY)
    $cleanTxtBtn.Size = New-Object System.Drawing.Size($btnW, $btnH)
    $cleanTxtBtn.FlatStyle = "Flat"
    $cleanTxtBtn.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(0, 210, 110)
    $cleanTxtBtn.BackColor = $darkBg
    $cleanTxtBtn.ForeColor = [System.Drawing.Color]::FromArgb(0, 210, 110)
    $cleanTxtBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $cleanTxtBtn.Add_Click({
        $cnt = ($script:ClipHistory | Where-Object { $_.Type -eq "text" }).Count
        if ($cnt -eq 0) { return }
        $r = [System.Windows.Forms.MessageBox]::Show(
            "Delete all text entries? ($cnt items)",
            "Clean Up Text", "YesNo", "Warning")
        if ($r -eq "Yes") {
            $newH = New-Object System.Collections.ArrayList
            foreach ($e in $script:ClipHistory) { if ($e.Type -ne "text") { [void]$newH.Add($e) } }
            $script:ClipHistory = $newH
            Save-History; Refresh-List
        }
    })
    $f.Controls.Add($cleanTxtBtn)

    $script:TxtStatsLabel = New-Object System.Windows.Forms.Label
    $script:TxtStatsLabel.Text = ""
    $script:TxtStatsLabel.Location = New-Object System.Drawing.Point(15, $statsY)
    $script:TxtStatsLabel.Size = New-Object System.Drawing.Size($btnW, 16)
    $script:TxtStatsLabel.ForeColor = [System.Drawing.Color]::FromArgb(140, 140, 140)
    $script:TxtStatsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 7)
    $script:TxtStatsLabel.TextAlign = "MiddleCenter"
    $f.Controls.Add($script:TxtStatsLabel)

    $cleanImgBtn = New-Object System.Windows.Forms.Button
    $cleanImgBtn.Text = "Clean images"
    $cleanImgBtn.Location = New-Object System.Drawing.Point((15 + $btnW + $gap), $btnY)
    $cleanImgBtn.Size = New-Object System.Drawing.Size(($btnW + 10), $btnH)
    $cleanImgBtn.FlatStyle = "Flat"
    $cleanImgBtn.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(60, 160, 255)
    $cleanImgBtn.BackColor = $darkBg
    $cleanImgBtn.ForeColor = [System.Drawing.Color]::FromArgb(60, 160, 255)
    $cleanImgBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $cleanImgBtn.Add_Click({
        $cnt = 0; $sz = [long]0
        if (Test-Path $script:ImageFolder) {
            Get-ChildItem $script:ImageFolder -File -ErrorAction SilentlyContinue | ForEach-Object {
                $cnt++; $sz += $_.Length
            }
        }
        if ($cnt -eq 0) { return }
        $szStr = Format-FileSize $sz
        $r = [System.Windows.Forms.MessageBox]::Show(
            "Delete all saved images?`n($cnt files, $szStr will be freed)",
            "Clean Up Images", "YesNo", "Warning")
        if ($r -eq "Yes") {
            Get-ChildItem $script:ImageFolder -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
            $newH = New-Object System.Collections.ArrayList
            foreach ($e in $script:ClipHistory) { if ($e.Type -ne "image") { [void]$newH.Add($e) } }
            $script:ClipHistory = $newH
            Save-History; Refresh-List
        }
    })
    $f.Controls.Add($cleanImgBtn)

    $script:ImgStatsLabel = New-Object System.Windows.Forms.Label
    $script:ImgStatsLabel.Text = ""
    $script:ImgStatsLabel.Location = New-Object System.Drawing.Point((15 + $btnW + $gap), $statsY)
    $script:ImgStatsLabel.Size = New-Object System.Drawing.Size(($btnW + 10), 16)
    $script:ImgStatsLabel.ForeColor = [System.Drawing.Color]::FromArgb(140, 140, 140)
    $script:ImgStatsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 7)
    $script:ImgStatsLabel.TextAlign = "MiddleCenter"
    $f.Controls.Add($script:ImgStatsLabel)

    $clearAllBtn = New-Object System.Windows.Forms.Button
    $clearAllBtn.Text = "Clear All"
    $clearAllBtn.Location = New-Object System.Drawing.Point((15 + $btnW + $gap + $btnW + 10 + $gap), $btnY)
    $clearAllBtn.Size = New-Object System.Drawing.Size($btnW, $btnH)
    $clearAllBtn.FlatStyle = "Flat"
    $clearAllBtn.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(255, 68, 68)
    $clearAllBtn.BackColor = $darkBg
    $clearAllBtn.ForeColor = [System.Drawing.Color]::FromArgb(255, 68, 68)
    $clearAllBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $clearAllBtn.Add_Click({
        $r = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to delete all clipboard history?",
            "Clear History", "YesNo", "Warning")
        if ($r -eq "Yes") {
            foreach ($e in $script:ClipHistory) {
                if ($e.Type -eq "image" -and $e.Image -and (Test-Path $e.Image)) {
                    Remove-Item $e.Image -Force -ErrorAction SilentlyContinue
                }
            }
            $script:ClipHistory.Clear()
            if (Test-Path $script:HistoryFile) { Remove-Item $script:HistoryFile -Force }
            Refresh-List
        }
    })
    $f.Controls.Add($clearAllBtn)

    $script:AllStatsLabel = New-Object System.Windows.Forms.Label
    $script:AllStatsLabel.Text = ""
    $script:AllStatsLabel.Location = New-Object System.Drawing.Point((15 + $btnW + $gap + $btnW + 10 + $gap), $statsY)
    $script:AllStatsLabel.Size = New-Object System.Drawing.Size($btnW, 16)
    $script:AllStatsLabel.ForeColor = [System.Drawing.Color]::FromArgb(140, 140, 140)
    $script:AllStatsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 7)
    $script:AllStatsLabel.TextAlign = "MiddleCenter"
    $f.Controls.Add($script:AllStatsLabel)

    Update-StorageStats

    # --- Form close -> hide ---
    $f.Add_FormClosing({
        param($sender, $eventArgs)
        if ($eventArgs.CloseReason -eq "UserClosing") {
            $eventArgs.Cancel = $true
            $f.Hide()
            $script:IsVisible = $false
            Close-Preview
        }
    })



    # --- Clipboard monitor timer ---
    $clipTimer = New-Object System.Windows.Forms.Timer
    $clipTimer.Interval = 500
    $clipTimer.Add_Tick({
        if ($global:SkipClipCheck) { return }
        try {
            if ([System.Windows.Forms.Clipboard]::ContainsImage()) {
                $img = [System.Windows.Forms.Clipboard]::GetImage()
                if ($null -ne $img) {
                    # Content-based fingerprint using sampled pixels
                    $fp = Get-ImageFingerprint $img
                    if ($fp -eq $global:LastImageFingerprint) {
                        $img.Dispose()
                        return
                    }
                    $global:LastImageFingerprint = $fp

                    $now = [DateTime]::Now
                    $ts = $now.ToString("yyyyMMdd_HHmmss")
                    $imgFile = Join-Path $script:ImageFolder ($ts + ".png")
                    $img.Save($imgFile, [System.Drawing.Imaging.ImageFormat]::Png)
                    $dims = [string]$img.Width + "x" + [string]$img.Height
                    $img.Dispose()

                    $newEntry = @{ Type="image"; Time=$now.ToString("MM/dd/yyyy HH:mm"); Text="[IMG] Screenshot $dims"; Image=$imgFile; Dims=$dims }
                    [void]$script:ClipHistory.Insert(0, $newEntry)
                    Enforce-MaxHistory
                    Save-History
                    if ($script:IsVisible) { Refresh-List }
                }
            }
            elseif ([System.Windows.Forms.Clipboard]::ContainsText()) {
                $text = [System.Windows.Forms.Clipboard]::GetText()
                if ([string]::IsNullOrEmpty($text) -or $text -eq $global:LastClipText) { return }
                if ($script:ClipHistory.Count -gt 0 -and $script:ClipHistory[0].Type -eq "text" -and $script:ClipHistory[0].Text -eq $text) { return }
                $global:LastClipText = $text
                $newEntry = @{ Type="text"; Time=[DateTime]::Now.ToString("MM/dd/yyyy HH:mm"); Text=$text; Image=""; Dims="" }
                [void]$script:ClipHistory.Insert(0, $newEntry)
                Enforce-MaxHistory
                Save-History
                if ($script:IsVisible) { Refresh-List }
            }
        }
        catch { }
    })
    $clipTimer.Start()

    # --- Initial list ---
    Refresh-List

    # --- Hotkey listener ---
    $hotkeyForm = New-Object HotkeyForm
    $hotkeyForm.Add_HotkeyPressed({ Toggle-Gui })
    $hotkeyForm.RegisterGlobalHotkey()

    # --- Tray icon ---
    $notifyIcon = New-Object System.Windows.Forms.NotifyIcon
    # Try to get clipboard icon from shell32.dll (index 260 = clipboard)
    $trayIcon = [System.Drawing.SystemIcons]::Application
    try {
        $lgIcons = New-Object IntPtr[] 1
        $smIcons = New-Object IntPtr[] 1
        [NativeMethods]::ExtractIconEx("shell32.dll", 260, $lgIcons, $smIcons, 1) | Out-Null
        if ($smIcons[0] -ne [IntPtr]::Zero) {
            $trayIcon = [System.Drawing.Icon]::FromHandle($smIcons[0])
        }
    } catch { }
    $notifyIcon.Icon = $trayIcon
    $notifyIcon.Text = "Clipboard Manager (Ctrl+Shift+H)"
    $notifyIcon.Visible = $true

    $trayMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $openItem = $trayMenu.Items.Add("Open  (Ctrl+Shift+H)")
    $openItem.Add_Click({ Toggle-Gui })
    [void]$trayMenu.Items.Add("-")
    $exitItem = $trayMenu.Items.Add("Exit")
    $exitItem.Add_Click({
        $clipTimer.Stop()
        $hotkeyForm.UnregisterGlobalHotkey()
        $notifyIcon.Visible = $false
        $notifyIcon.Dispose()
        # preview removed
        $f.Close()
        $f.Dispose()
        $hotkeyForm.Close()
        [System.Windows.Forms.Application]::Exit()
    })
    $notifyIcon.ContextMenuStrip = $trayMenu
    $notifyIcon.Add_DoubleClick({ Toggle-Gui })

    # Run message loop
    $hotkeyForm.Show()
    [System.Windows.Forms.Application]::Run($hotkeyForm)
}

# ============================================================
Build-MainForm
