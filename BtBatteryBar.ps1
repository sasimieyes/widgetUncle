# ============================================================
# BtBatteryBar.ps1
# 블루투스 장치별 배터리 잔량을 작업표시줄 위에 표시하는 바 위젯
#
# 방식: 최상위(TopMost) 무테두리 투명 창을 작업표시줄 영역 좌표에 겹쳐 배치.
#   Windows 11 24H2/25H2 의 XAML 작업표시줄은 SetParent 자식 부착을
#   즉시 되돌리므로 (호출은 성공하지만 부모가 원복됨) 오버레이가 유일한 방법.
# 특징:
#   - 배경 투명 (TransparencyKey, 키 색은 작업표시줄 색 근사치로 잔티 최소화)
#   - 알림 영역(시계) 왼쪽에 우측 정렬, 드래그(글자/아이콘 위)로 위치 조정
#   - 표시 방식: 숫자(85%) 또는 배터리 모양 아이콘
#   - 아이콘 색: 컬러(기본 녹색/20% 이하 노랑/10% 이하 빨강, 기본색 변경 가능)
#                또는 그레이스케일(테마 단색)
#   - 현재 연결된 장치만 표시 (System.Devices.Aep.IsConnected)
#   - 장치별 표시 on/off, 표시 이름(별칭) 변경 — 설정 창 (Solarized 스타일)
#   - 1초 감시: 작업표시줄 추적, 전체화면 앱 감지 숨김, 최상위 유지
#   - WS_EX_NOACTIVATE: 클릭해도 다른 앱의 포커스를 뺏지 않음
# 사용법: wscript start_bar.vbs (창 없이 실행)
#   테스트: powershell -STA -File BtBatteryBar.ps1 -Test
#   설정창 단독 테스트: powershell -STA -File BtBatteryBar.ps1 -SettingsTest
# ============================================================
param(
    [switch]$Test,
    [switch]$SettingsTest
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public class BarApi {
    [DllImport("user32.dll", SetLastError = true)] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    [DllImport("user32.dll", SetLastError = true)] public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    [DllImport("user32.dll")] public static extern int GetWindowLong(IntPtr hWnd, int nIndex);
    [DllImport("user32.dll")] public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
}
'@

[void][BarApi]::SetProcessDPIAware()

$script:BaseDir      = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:SettingsPath = Join-Path $script:BaseDir 'settings_bar.json'
$script:LogPath      = Join-Path $script:BaseDir 'widget.log'
$script:StartupLnk   = Join-Path ([Environment]::GetFolderPath('Startup')) 'BtBatteryBar.lnk'
$script:BatteryKey   = '{104EA319-6EE2-4701-BD47-8DDBF425BBE5} 2'  # DEVPKEY_Bluetooth_Battery
$script:ConnectedKey = '{83DA6326-97A6-4088-9453-A1923F573B29} 15' # System.Devices.Aep.IsConnected (BT 디바이스 노드에 미러링됨)

# DPI 배율 (설정창 치수에 사용)
$script:DpiScale = 1.0
try {
    $g0 = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
    $script:DpiScale = $g0.DpiX / 96.0
    $g0.Dispose()
} catch { }
function S { param([double]$v) return [int][Math]::Round($v * $script:DpiScale) }

# 전체화면 판정에서 제외할 셸 프로세스 (시작 메뉴/검색/잠금 등)
$script:ShellProcNames = @('explorer', 'StartMenuExperienceHost', 'SearchHost', 'SearchApp', 'ShellExperienceHost', 'ShellHost', 'LockApp', 'dwm')

# 설정창 테마: Solarized Light + 부트스트랩 느낌
$script:UI = @{
    Base3  = [System.Drawing.ColorTranslator]::FromHtml('#FDF6E3')   # 폼 배경
    Base2  = [System.Drawing.ColorTranslator]::FromHtml('#EEE8D5')   # 카드/보조 버튼
    Base1  = [System.Drawing.ColorTranslator]::FromHtml('#93A1A1')   # 테두리
    Base01 = [System.Drawing.ColorTranslator]::FromHtml('#586E75')   # 본문 텍스트
    Base02 = [System.Drawing.ColorTranslator]::FromHtml('#073642')   # 진한 텍스트
    Blue   = [System.Drawing.ColorTranslator]::FromHtml('#268BD2')   # primary
    Sel    = [System.Drawing.ColorTranslator]::FromHtml('#D5E8F5')   # 그리드 선택
}

# ------------------------------------------------------------ 설정/로그
$script:Settings = @{
    RefreshSec     = 60
    BarOffsetRight = 0           # 알림 영역 왼쪽 기준점에서 왼쪽으로 띄울 거리(px)
    DisplayMode    = 'icon'      # 'number' | 'icon' | 'iconNumber'
    HiddenDevices  = @()
    ConnectedOnly  = $true       # 현재 연결된 장치만 표시
    Aliases        = @{}         # 원본 장치명 -> 표시 이름
    IconColorMode  = 'color'     # 'color' | 'gray'
    NormalColor    = '#4CAF50'   # 컬러 모드의 정상 잔량 색 (기본 녹색)
    WarnColor      = '#FFCD3C'   # 컬러 모드의 20% 이하 경고 색
    IconTextColor  = '#FFFFFF'   # 배터리 아이콘 안 숫자 색 (게이지 위)
    InvertIconText = $true       # 게이지/빈 영역 경계에서 숫자 색 반전
}

function Read-Settings {
    if (-not (Test-Path $script:SettingsPath)) { return }
    try {
        $j = Get-Content $script:SettingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($j.RefreshSec -ge 10) { $script:Settings.RefreshSec = [int]$j.RefreshSec }
        if ($null -ne $j.BarOffsetRight -and $j.BarOffsetRight -ge 0) { $script:Settings.BarOffsetRight = [int]$j.BarOffsetRight }
        if ($j.DisplayMode -eq 'icon' -or $j.DisplayMode -eq 'number' -or $j.DisplayMode -eq 'iconNumber') { $script:Settings.DisplayMode = [string]$j.DisplayMode }
        if ($null -ne $j.HiddenDevices) { $script:Settings.HiddenDevices = @($j.HiddenDevices | ForEach-Object { [string]$_ }) }
        if ($null -ne $j.ConnectedOnly) { $script:Settings.ConnectedOnly = [bool]$j.ConnectedOnly }
        if ($null -ne $j.Aliases) {
            $a = @{}
            foreach ($prop in $j.Aliases.PSObject.Properties) {
                if (-not [string]::IsNullOrWhiteSpace($prop.Value)) { $a[[string]$prop.Name] = [string]$prop.Value }
            }
            $script:Settings.Aliases = $a
        }
        if ($j.IconColorMode -eq 'color' -or $j.IconColorMode -eq 'gray') { $script:Settings.IconColorMode = [string]$j.IconColorMode }
        if ($j.NormalColor -match '^#[0-9A-Fa-f]{6}$') { $script:Settings.NormalColor = [string]$j.NormalColor }
        if ($j.WarnColor -match '^#[0-9A-Fa-f]{6}$') { $script:Settings.WarnColor = [string]$j.WarnColor }
        if ($j.IconTextColor -match '^#[0-9A-Fa-f]{6}$') { $script:Settings.IconTextColor = [string]$j.IconTextColor }
        if ($null -ne $j.InvertIconText) { $script:Settings.InvertIconText = [bool]$j.InvertIconText }
    } catch { }
}

function Save-Settings {
    try {
        $obj = [PSCustomObject]@{
            RefreshSec     = $script:Settings.RefreshSec
            BarOffsetRight = $script:Settings.BarOffsetRight
            DisplayMode    = $script:Settings.DisplayMode
            HiddenDevices  = @($script:Settings.HiddenDevices)
            ConnectedOnly  = $script:Settings.ConnectedOnly
            Aliases        = $script:Settings.Aliases
            IconColorMode  = $script:Settings.IconColorMode
            NormalColor    = $script:Settings.NormalColor
            WarnColor      = $script:Settings.WarnColor
            IconTextColor  = $script:Settings.IconTextColor
            InvertIconText = $script:Settings.InvertIconText
        }
        [System.IO.File]::WriteAllText($script:SettingsPath, ($obj | ConvertTo-Json), (New-Object System.Text.UTF8Encoding $true))
    } catch { }
}

function Write-Log {
    param([string]$Message)
    try {
        $line = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
        Add-Content -Path $script:LogPath -Value $line -Encoding UTF8
    } catch { }
}

# ------------------------------------------------------------ 배터리 조회
# 배터리(DEVPKEY_Bluetooth_Battery)와 실제 연결 여부(System.Devices.Aep.IsConnected)를
# 한 번의 GetDeviceProperties 호출로 함께 조회한다.
function Get-BtBatteryList {
    $byName = @{}
    try {
        $devices = Get-CimInstance -ClassName Win32_PnPEntity -Filter "PNPClass='Bluetooth'" -ErrorAction Stop
    } catch {
        return @()
    }
    foreach ($dev in $devices) {
        if ([string]::IsNullOrWhiteSpace($dev.Name)) { continue }
        if ($dev.Status -ne 'OK') { continue }
        $pct = $null
        $conn = $null
        try {
            $resp = Invoke-CimMethod -InputObject $dev -MethodName GetDeviceProperties `
                -Arguments @{ devicePropertyKeys = @($script:BatteryKey, $script:ConnectedKey) } -ErrorAction Stop
            foreach ($prop in @($resp.deviceProperties)) {
                if ($null -eq $prop.Data) { continue }
                if ($prop.KeyName -eq $script:BatteryKey) { $pct = [int]$prop.Data }
                elseif ($prop.KeyName -eq $script:ConnectedKey) { $conn = [bool]$prop.Data }
            }
        } catch { continue }
        if ($null -eq $pct -or $pct -lt 0 -or $pct -gt 100) { continue }
        if (-not $byName.ContainsKey($dev.Name)) {
            $byName[$dev.Name] = New-Object System.Collections.ArrayList
        }
        [void]$byName[$dev.Name].Add(@{ Pct = $pct; Conn = $conn })
    }
    $list = foreach ($name in ($byName.Keys | Sort-Object)) {
        $entries = @($byName[$name])

        # 연결된 인스턴스가 있으면 그 값만 사용 (이전 연결의 잔존 배터리 값 배제)
        $connEntries = @($entries | Where-Object { $_.Conn -eq $true })
        $connected = $false
        $use = $entries
        if ($connEntries.Count -gt 0) {
            $connected = $true
            $use = $connEntries
        } else {
            # 연결 속성을 아예 제공하지 않는 장치는 판단 불가 → 표시 유지
            $known = @($entries | Where-Object { $null -ne $_.Conn })
            if ($known.Count -eq 0) { $connected = $true }
        }

        # 표시 이름: 사용자 별칭 우선, 없으면 자동 약칭
        $shortName = $null
        if ($script:Settings.Aliases.ContainsKey($name)) {
            $a = [string]$script:Settings.Aliases[$name]
            if (-not [string]::IsNullOrWhiteSpace($a)) { $shortName = $a.Trim() }
        }
        if ($null -eq $shortName) { $shortName = Get-ShortName -Name $name }

        $vals = @($use | ForEach-Object { $_.Pct } | Sort-Object)
        $uniq = @($vals | Select-Object -Unique)
        [PSCustomObject]@{
            Name      = $name
            Short     = $shortName
            Percent   = $vals[0]
            Detail    = (($uniq | ForEach-Object { '{0}%' -f $_ }) -join ' / ')
            Connected = $connected
        }
    }
    return @($list)
}

# 장치명 약칭: 마지막 단어(모델명) 사용, 너무 짧으면 앞 단어 결합
function Get-ShortName {
    param([string]$Name)
    $tokens = @($Name -split '\s+' | Where-Object { $_ })
    if ($tokens.Count -eq 0) { return $Name }
    $short = $tokens[-1]
    if ($tokens.Count -ge 2 -and $short.Length -le 3) {
        $short = '{0} {1}' -f $tokens[-2], $short
    }
    if ($short.Length -gt 12) { $short = $short.Substring(0, 11) + '..' }
    return $short
}

function Get-TaskbarIsLight {
    try {
        $v = Get-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize' `
            -Name SystemUsesLightTheme -ErrorAction Stop
        return ($v.SystemUsesLightTheme -eq 1)
    } catch { return $false }
}

# 작업표시줄의 실제 색을 1픽셀 샘플링 (위젯 배경 위장용)
# 샘플 지점: 작업표시줄 상단에서 3px 아래 (위젯은 +7px 부터라 겹치지 않음)
function Get-TaskbarSampleColor {
    param($TrayRect)
    if ($null -eq $TrayRect) { return $null }
    try {
        $sx = [int]((Get-AnchorRight -TrayRect $TrayRect) - 60)
        if ($sx -lt ($TrayRect.Left + 5)) { $sx = $TrayRect.Left + 5 }
        $sy = $TrayRect.Top + 3
        $bmp = New-Object System.Drawing.Bitmap(1, 1)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.CopyFromScreen($sx, $sy, 0, 0, (New-Object System.Drawing.Size(1, 1)))
        $g.Dispose()
        $c = $bmp.GetPixel(0, 0)
        $bmp.Dispose()
        return [System.Drawing.Color]::FromArgb(255, $c.R, $c.G, $c.B)
    } catch { return $null }
}

function Get-Palette {
    param([bool]$Light)
    if ($Light) {
        @{
            Key    = [System.Drawing.Color]::FromArgb(238, 238, 238)  # 투명 키 (작업표시줄 색 근사)
            Name   = [System.Drawing.Color]::FromArgb(95, 95, 95)
            Normal = [System.Drawing.Color]::FromArgb(25, 25, 25)
            Warn   = [System.Drawing.Color]::FromArgb(190, 140, 0)    # 노랑 (라이트용 진한 톤)
            Crit   = [System.Drawing.Color]::FromArgb(210, 50, 50)    # 빨강
        }
    } else {
        @{
            Key    = [System.Drawing.Color]::FromArgb(28, 28, 28)
            Name   = [System.Drawing.Color]::FromArgb(165, 165, 165)
            Normal = [System.Drawing.Color]::FromArgb(240, 240, 240)
            Warn   = [System.Drawing.Color]::FromArgb(255, 205, 60)
            Crit   = [System.Drawing.Color]::FromArgb(255, 82, 82)
        }
    }
}

function ConvertTo-ColorSafe {
    param([string]$Hex, [System.Drawing.Color]$Fallback)
    try {
        if ([string]::IsNullOrWhiteSpace($Hex)) { return $Fallback }
        return [System.Drawing.ColorTranslator]::FromHtml($Hex)
    } catch { return $Fallback }
}

# 잔량 → 상태색
#  컬러 모드     : 정상=사용자 지정색(기본 녹색), 20% 이하 경고색, 10% 이하 빨강
#  그레이스케일  : 테마 단색 (흰/검)
function Get-LevelColor {
    param([int]$Percent, $Pal)
    if ($script:Settings.IconColorMode -eq 'gray') { return $Pal.Normal }
    if ($Percent -le 10) { return $Pal.Crit }
    if ($Percent -le 20) { return (ConvertTo-ColorSafe $script:Settings.WarnColor $Pal.Warn) }
    return (ConvertTo-ColorSafe $script:Settings.NormalColor ([System.Drawing.Color]::FromArgb(76, 175, 80)))
}

function Get-IconNumberColor {
    return (ConvertTo-ColorSafe $script:Settings.IconTextColor ([System.Drawing.Color]::White))
}

# ------------------------------------------------------------ 배터리 아이콘 렌더링
# 가로형 배터리: 테두리 + 오른쪽 양극 캡 + 잔량만큼 채움
function New-BatteryImage {
    param([int]$Percent, $Pal, [int]$W, [int]$H, [switch]$ShowPercent)
    $bmp = New-Object System.Drawing.Bitmap($W, $H, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb))
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
        $g.Clear([System.Drawing.Color]::Transparent)

        $levelColor = Get-LevelColor -Percent $Percent -Pal $Pal
        $penW   = [Math]::Max(1.5, [Math]::Round($H * 0.09, 1))
        $capW   = [Math]::Max(2.0, [Math]::Round($W * 0.07, 1))
        $capGap = [Math]::Max(1.0, [Math]::Round($penW * 0.45, 1))
        $radius = [Math]::Max(2.5, [Math]::Round($H * 0.19, 1))

        $bodyW = $W - $capW - $capGap - $penW
        $bodyH = $H - $penW
        $bx = $penW / 2
        $by = ($H - $bodyH) / 2
        $bw = $bodyW
        $bh = $bodyH

        # 본체 외곽 (둥근 사각형)
        $path = New-Object System.Drawing.Drawing2D.GraphicsPath
        $d = $radius * 2
        $path.AddArc([float]$bx, [float]$by, [float]$d, [float]$d, 180, 90)
        $path.AddArc([float]($bx + $bw - $d), [float]$by, [float]$d, [float]$d, 270, 90)
        $path.AddArc([float]($bx + $bw - $d), [float]($by + $bh - $d), [float]$d, [float]$d, 0, 90)
        $path.AddArc([float]$bx, [float]($by + $bh - $d), [float]$d, [float]$d, 90, 90)
        $path.CloseFigure()
        $pen = New-Object System.Drawing.Pen($Pal.Name, [float]$penW)
        $g.DrawPath($pen, $path)

        # 양극 캡
        $capH = $bh * 0.46
        $capBrush = New-Object System.Drawing.SolidBrush($Pal.Name)
        $capRect = New-Object System.Drawing.RectangleF([float]($bx + $bw + $capGap), [float]($by + ($bh - $capH) / 2), [float]$capW, [float]$capH)
        $g.FillRectangle($capBrush, $capRect)

        # 잔량 채움
        $innerGap = [Math]::Max(2.0, [Math]::Round($H * 0.12, 1))
        $ix = $bx + ($penW / 2) + $innerGap
        $iy = $by + ($penW / 2) + $innerGap
        $iw = $bw - $penW - ($innerGap * 2)
        $ih = $bh - $penW - ($innerGap * 2)
        if ($iw -lt 1) { $iw = 1 }
        if ($ih -lt 1) { $ih = 1 }
        $fillW = [float]($iw * $Percent / 100.0)
        if ($Percent -gt 0 -and $fillW -lt [Math]::Min(2.0, $iw)) { $fillW = [Math]::Min(2.0, $iw) }
        if ($fillW -gt 0) {
            $fillBrush = New-Object System.Drawing.SolidBrush($levelColor)
            $fr = [Math]::Max(1.0, $radius - (($innerGap + ($penW / 2)) / 2))
            if ($fr -gt ($ih / 2)) { $fr = $ih / 2 }
            $fd = $fr * 2
            $fx = $ix
            $fy = $iy
            $fh = $ih
            if ($fillW -le $fd) {
                $g.FillEllipse($fillBrush, [float]$fx, [float]$fy, [float][Math]::Max(2, $fillW), [float]$fh)
            } else {
                $fpath = New-Object System.Drawing.Drawing2D.GraphicsPath
                $fpath.AddArc([float]$fx, [float]$fy, [float]$fd, [float]$fd, 180, 90)
                $fpath.AddArc([float]($fx + $fillW - $fd), [float]$fy, [float]$fd, [float]$fd, 270, 90)
                $fpath.AddArc([float]($fx + $fillW - $fd), [float]($fy + $fh - $fd), [float]$fd, [float]$fd, 0, 90)
                $fpath.AddArc([float]$fx, [float]($fy + $fh - $fd), [float]$fd, [float]$fd, 90, 90)
                $fpath.CloseFigure()
                $g.FillPath($fillBrush, $fpath)
                $fpath.Dispose()
            }
            $fillBrush.Dispose()
        }

        if ($ShowPercent) {
            $textColor = Get-IconNumberColor
            $emptyTextColor = $Pal.Normal
            $fontSize = [Math]::Max(8, [int]($H * 0.50))
            $font = New-Object System.Drawing.Font('Segoe UI', $fontSize, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
            $fmt = New-Object System.Drawing.StringFormat
            $fmt.Alignment = [System.Drawing.StringAlignment]::Center
            $fmt.LineAlignment = [System.Drawing.StringAlignment]::Center
            $textYOffset = [Math]::Max(0.5, [Math]::Round($H * 0.03, 1))
            $textRect = New-Object System.Drawing.RectangleF([float]$bx, [float]($by + $textYOffset), [float]$bw, [float]$bh)
            if ($script:Settings.InvertIconText) {
                $emptyBrush = New-Object System.Drawing.SolidBrush($emptyTextColor)
                $g.DrawString([string]$Percent, $font, $emptyBrush, $textRect, $fmt)
                $emptyBrush.Dispose()

                if ($fillW -gt 0) {
                    $oldClip = $g.Clip
                    $filledClip = New-Object System.Drawing.RectangleF([float]$ix, [float]$iy, [float]$fillW, [float]$ih)
                    $g.SetClip($filledClip)
                    $textBrush = New-Object System.Drawing.SolidBrush($textColor)
                    $g.DrawString([string]$Percent, $font, $textBrush, $textRect, $fmt)
                    $textBrush.Dispose()
                    $g.Clip = $oldClip
                    $oldClip.Dispose()
                }
            } else {
                $textBrush = New-Object System.Drawing.SolidBrush($textColor)
                $g.DrawString([string]$Percent, $font, $textBrush, $textRect, $fmt)
                $textBrush.Dispose()
            }
            $fmt.Dispose(); $font.Dispose()
        }

        $pen.Dispose(); $path.Dispose(); $capBrush.Dispose()
        return $bmp
    } finally {
        $g.Dispose()
    }
}

# ------------------------------------------------------------ 작업표시줄 위치 (우측 정렬)
$script:TrayHwnd    = [IntPtr]::Zero
$script:LastTrayKey = ''

function Get-TrayRect {
    if (-not [BarApi]::IsWindow($script:TrayHwnd)) {
        $script:TrayHwnd = [BarApi]::FindWindow('Shell_TrayWnd', [NullString]::Value)
    }
    if ($script:TrayHwnd -ne [IntPtr]::Zero) {
        $r = New-Object 'BarApi+RECT'
        if ([BarApi]::GetWindowRect($script:TrayHwnd, [ref]$r)) {
            if (($r.Bottom - $r.Top) -ge 20 -and ($r.Bottom - $r.Top) -le 200) { return $r }
        }
    }
    return $null
}

# 우측 기준점: 알림 영역(TrayNotifyWnd, 시계/숨김아이콘) 바로 왼쪽
function Get-AnchorRight {
    param($TrayRect)
    if ($script:TrayHwnd -ne [IntPtr]::Zero) {
        $notify = [BarApi]::FindWindowEx($script:TrayHwnd, [IntPtr]::Zero, 'TrayNotifyWnd', [NullString]::Value)
        if ($notify -ne [IntPtr]::Zero) {
            $nr = New-Object 'BarApi+RECT'
            if ([BarApi]::GetWindowRect($notify, [ref]$nr)) {
                if ($nr.Left -gt $TrayRect.Left -and $nr.Left -le $TrayRect.Right) {
                    return $nr.Left - 10
                }
            }
        }
    }
    return $TrayRect.Right - 10
}

function Position-Bar {
    $r = Get-TrayRect
    if ($null -ne $r) {
        $tbH = $r.Bottom - $r.Top
        $barH = [Math]::Max(26, $tbH - 14)
        if ($script:Form.Height -ne $barH) { $script:Form.Height = $barH }
        $anchor = Get-AnchorRight -TrayRect $r
        $left = $anchor - [Math]::Max(0, $script:Settings.BarOffsetRight) - $script:Form.Width
        if ($left -lt $r.Left) { $left = $r.Left }
        $script:Form.Top  = $r.Top + [int](($tbH - $barH) / 2)
        $script:Form.Left = $left
        $script:LastTrayKey = '{0},{1},{2},{3},{4}' -f $r.Left, $r.Top, $r.Right, $r.Bottom, $anchor
    } else {
        # 작업표시줄 창을 찾지 못하면 주 화면 우측 하단에 표시
        $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        $script:Form.Top  = $wa.Bottom - $script:Form.Height - 6
        $script:Form.Left = $wa.Right - $script:Form.Width - 8 - [Math]::Max(0, $script:Settings.BarOffsetRight)
        $script:LastTrayKey = ''
    }
}

function Force-TopMost {
    if ($script:Form.IsHandleCreated -and -not $script:Form.IsDisposed) {
        # HWND_TOPMOST(-1), SWP_NOSIZE|SWP_NOMOVE|SWP_NOACTIVATE
        [void][BarApi]::SetWindowPos($script:Form.Handle, [IntPtr](-1), 0, 0, 0, 0, 0x0013)
    }
}

# 클릭해도 포커스를 뺏지 않도록 + Alt-Tab 목록에서 제외
function Set-BarExStyle {
    $GWL_EXSTYLE = -20
    $ex = [BarApi]::GetWindowLong($script:Form.Handle, $GWL_EXSTYLE)
    $ex = $ex -bor 0x08000000 -bor 0x00000080   # WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW
    [void][BarApi]::SetWindowLong($script:Form.Handle, $GWL_EXSTYLE, $ex)
}

# 전체화면 앱이 앞에 있으면 바를 숨김 (게임/동영상 방해 방지)
function Test-FullscreenForeground {
    $fg = [BarApi]::GetForegroundWindow()
    if ($fg -eq [IntPtr]::Zero) { return $false }

    $procId = [uint32]0
    [void][BarApi]::GetWindowThreadProcessId($fg, [ref]$procId)
    if ($procId -eq $PID) { return $false }
    try {
        $pname = [System.Diagnostics.Process]::GetProcessById([int]$procId).ProcessName
        if ($script:ShellProcNames -contains $pname) { return $false }
    } catch { }

    # 클릭 통과형 투명 오버레이(WS_EX_TRANSPARENT)나 포커스를 받지 않는 창(WS_EX_NOACTIVATE)은
    # 화면 전체 크기라도 전체화면 앱이 아님 (예: 모니터 유틸리티/캡처 도구의 투명 오버레이)
    $exStyle = [BarApi]::GetWindowLong($fg, -20)
    if (($exStyle -band 0x00000020) -ne 0) { return $false }   # WS_EX_TRANSPARENT
    if (($exStyle -band 0x08000000) -ne 0) { return $false }   # WS_EX_NOACTIVATE

    $sb = New-Object System.Text.StringBuilder 256
    [void][BarApi]::GetClassName($fg, $sb, 256)
    $cls = $sb.ToString()
    if ($cls -eq 'Progman' -or $cls -eq 'WorkerW' -or $cls -eq 'Shell_TrayWnd' -or $cls -eq 'Windows.UI.Core.CoreWindow' -or $cls -eq 'XamlExplorerHostIslandWindow') { return $false }

    $r = New-Object 'BarApi+RECT'
    if (-not [BarApi]::GetWindowRect($fg, [ref]$r)) { return $false }
    $cx = [int](($r.Left + $r.Right) / 2)
    $cy = [int](($r.Top + $r.Bottom) / 2)
    $scr = [System.Windows.Forms.Screen]::FromPoint((New-Object System.Drawing.Point($cx, $cy)))
    $b = $scr.Bounds
    return ($r.Left -le $b.Left -and $r.Top -le $b.Top -and $r.Right -ge $b.Right -and $r.Bottom -ge $b.Bottom)
}

# ------------------------------------------------------------ 드래그 이동
$script:DragOffX = -1

function Add-DragHandlers {
    param($Control)
    $Control.ContextMenuStrip = $script:Menu
    $Control.add_MouseDown({
        param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            $script:DragOffX = [System.Windows.Forms.Control]::MousePosition.X - $script:Form.Left
        }
    })
    $Control.add_MouseMove({
        param($s, $e)
        if ($script:DragOffX -ge 0 -and $e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            $newLeft = [System.Windows.Forms.Control]::MousePosition.X - $script:DragOffX
            $script:Form.Left = [Math]::Max(0, $newLeft)
        }
    })
    $Control.add_MouseUp({
        param($s, $e)
        if ($script:DragOffX -ge 0) {
            $script:DragOffX = -1
            $r = Get-TrayRect
            if ($null -ne $r) {
                $anchor = Get-AnchorRight -TrayRect $r
                $script:Settings.BarOffsetRight = [Math]::Max(0, $anchor - ($script:Form.Left + $script:Form.Width))
            }
            Save-Settings
            Position-Bar
        }
    })
    # 더블클릭 새로고침은 마우스 이벤트 시퀀스가 끝난 뒤 실행한다.
    # (이벤트 중간에 라벨을 재구성하면 MouseUp 이 소실되어 드래그 상태가 고착됨)
    $Control.add_DoubleClick({
        $script:DragOffX = -1
        if ($null -ne $script:OnceTimer) {
            $script:OnceTimer.Stop()
            $script:OnceTimer.Start()
        }
    })
}

# ------------------------------------------------------------ 라벨/아이콘 구성
$script:LastAll = @()   # 감지된 전체 장치 (설정창용)

function Clear-BarControls {
    foreach ($c in @($script:Form.Controls)) {
        $script:Form.Controls.Remove($c)
        if ($c -is [System.Windows.Forms.PictureBox] -and $null -ne $c.Image) {
            $img = $c.Image; $c.Image = $null; $img.Dispose()
        }
        $c.Dispose()
    }
}

function Build-Labels {
    param($List, [bool]$Light, [System.Drawing.Color]$HitBack)
    $pal = Get-Palette -Light $Light
    $form = $script:Form
    $form.SuspendLayout()
    Clear-BarControls

    # 위장 배경: 작업표시줄에서 샘플링한 실제 색으로 전체를 칠한다.
    # 시각적으로는 투명과 같지만, 바 사각형 전체가 마우스 입력을 받는다.
    # (TransparencyKey 색상키 방식은 글자 획 픽셀만 클릭이 잡혀 사용 불가)
    $form.BackColor = $HitBack

    $barH = $form.Height
    $nameSize = [Math]::Max(9, [int]($barH * 0.34))
    $valSize  = [Math]::Max(10, [int]($barH * 0.40))
    $nameFont = New-Object System.Drawing.Font('Segoe UI', $nameSize, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)
    $valFont  = New-Object System.Drawing.Font('Segoe UI', $valSize, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $displayMode = [string]$script:Settings.DisplayMode
    $iconMode = ($displayMode -eq 'icon' -or $displayMode -eq 'iconNumber')
    $iconHasNumber = ($displayMode -eq 'iconNumber')
    $iconW = if ($iconHasNumber) { [Math]::Max(34, [int]($barH * 0.88)) } else { [Math]::Max(22, [int]($barH * 0.62)) }
    $iconH = if ($iconHasNumber) { [Math]::Max(17, [int]($barH * 0.48)) } else { [Math]::Max(12, [int]($barH * 0.38)) }

    # 모든 컨트롤을 바 전체 높이로 만들고 간격을 패딩으로 흡수해
    # 바 영역에 클릭이 통과하는 빈틈이 없게 한다.
    $x = 0
    if ($List.Count -eq 0) {
        $msg = 'BT 배터리 장치 없음'
        if ($script:LastAll.Count -gt 0) { $msg = '표시할 장치 없음 (우클릭-설정)' }
        $lb = New-Object System.Windows.Forms.Label
        $lb.AutoSize = $false
        $lb.Font = $nameFont
        $lb.ForeColor = $pal.Name
        $lb.BackColor = $HitBack
        $lb.Text = $msg
        $lb.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        $lb.Padding = New-Object System.Windows.Forms.Padding(12, 0, 12, 0)
        $lb.Size = New-Object System.Drawing.Size($lb.PreferredWidth, $barH)
        $lb.Location = New-Object System.Drawing.Point($x, 0)
        $form.Controls.Add($lb)
        Add-DragHandlers -Control $lb
        $script:Tip.SetToolTip($lb, '우클릭 메뉴의 설정에서 표시할 장치를 선택할 수 있습니다')
        $x += $lb.Width
    } else {
        $first = $true
        foreach ($item in $List) {
            $tipText = '{0}  {1}' -f $item.Name, $item.Detail
            $leftPad = if ($first) { 12 } else { 14 }
            $first = $false

            $lbName = New-Object System.Windows.Forms.Label
            $lbName.AutoSize = $false
            $lbName.Font = $nameFont
            $lbName.ForeColor = $pal.Name
            $lbName.BackColor = $HitBack
            $lbName.Text = $item.Short
            $lbName.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
            $lbName.Padding = New-Object System.Windows.Forms.Padding($leftPad, 0, 4, 0)
            $lbName.Size = New-Object System.Drawing.Size($lbName.PreferredWidth, $barH)
            $lbName.Location = New-Object System.Drawing.Point($x, 0)
            $form.Controls.Add($lbName)
            $x += $lbName.Width
            $script:Tip.SetToolTip($lbName, $tipText)
            Add-DragHandlers -Control $lbName

            if ($iconMode) {
                $pb = New-Object System.Windows.Forms.PictureBox
                $pb.Size = New-Object System.Drawing.Size($iconW, $barH)
                $pb.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::CenterImage
                $pb.BackColor = $HitBack
                $pb.Image = New-BatteryImage -Percent $item.Percent -Pal $pal -W $iconW -H $iconH -ShowPercent:$iconHasNumber
                $pb.Location = New-Object System.Drawing.Point($x, 0)
                $form.Controls.Add($pb)
                $x += $pb.Width
                $script:Tip.SetToolTip($pb, $tipText)
                Add-DragHandlers -Control $pb
            } else {
                $lbVal = New-Object System.Windows.Forms.Label
                $lbVal.AutoSize = $false
                $lbVal.Font = $valFont
                $lbVal.ForeColor = (Get-LevelColor -Percent $item.Percent -Pal $pal)
                $lbVal.BackColor = $HitBack
                $lbVal.Text = '{0}%' -f $item.Percent
                $lbVal.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                $lbVal.Size = New-Object System.Drawing.Size($lbVal.PreferredWidth, $barH)
                $lbVal.Location = New-Object System.Drawing.Point($x, 0)
                $form.Controls.Add($lbVal)
                $x += $lbVal.Width
                $script:Tip.SetToolTip($lbVal, $tipText)
                Add-DragHandlers -Control $lbVal
            }
        }
    }

    $form.Width = $x + 12   # 오른쪽 여백 (폼 배경도 같은 색이라 히트 영역)
    $form.ResumeLayout()
}

# ------------------------------------------------------------ 갱신
function Update-Bar {
    param([switch]$Force)
    if (-not $Force) {
        if ($script:Menu.Visible) { return }
        if ($script:DragOffX -ge 0) { return }
    }

    # 작업표시줄 높이/배경색을 먼저 확정 (라벨이 바 전체 높이를 차지하므로)
    $r = Get-TrayRect
    $barH = 44
    if ($null -ne $r) { $barH = [Math]::Max(26, ($r.Bottom - $r.Top) - 14) }
    if ($script:Form.Height -ne $barH) { $script:Form.Height = $barH }

    $sample = Get-TaskbarSampleColor -TrayRect $r
    $light = Get-TaskbarIsLight
    if ($null -ne $sample) {
        $lum = 0.299 * $sample.R + 0.587 * $sample.G + 0.114 * $sample.B
        $light = ($lum -gt 127)
    }
    $hit = if ($null -ne $sample) { $sample } else { (Get-Palette -Light $light).Key }

    $all = Get-BtBatteryList
    $script:LastAll = $all
    $hidden = @($script:Settings.HiddenDevices)
    $visible = @($all | Where-Object { $hidden -notcontains $_.Name -and ($_.Connected -or -not $script:Settings.ConnectedOnly) })
    Build-Labels -List $visible -Light $light -HitBack $hit
    Position-Bar
    Force-TopMost
    Update-Menu
}

# 1초마다: 전체화면 감지 / 작업표시줄 이동 추적 / 최상위 유지
function Watch-Tick {
    if ($script:Form.IsDisposed) { return }
    if ($script:DragOffX -ge 0) {
        # MouseUp 소실로 드래그 상태가 고착되는 것을 자가 복구
        if ([System.Windows.Forms.Control]::MouseButtons -eq [System.Windows.Forms.MouseButtons]::None) {
            $script:DragOffX = -1
        } else {
            return
        }
    }
    if ($script:Menu.Visible) { return }

    $fs = Test-FullscreenForeground
    if ($fs) {
        if ($script:Form.Visible) { $script:Form.Visible = $false }
        return
    }
    if (-not $script:Form.Visible) {
        $script:Form.Visible = $true
        Position-Bar
        Force-TopMost
        return
    }

    $r = Get-TrayRect
    $key = ''
    if ($null -ne $r) {
        $anchor = Get-AnchorRight -TrayRect $r
        $key = '{0},{1},{2},{3},{4}' -f $r.Left, $r.Top, $r.Right, $r.Bottom, $anchor
    }
    if ($key -ne $script:LastTrayKey) { Position-Bar }

    # 작업표시줄/다른 창 클릭으로 가려지는 것을 즉시 복구
    Force-TopMost
}

# ------------------------------------------------------------ 설정 창 (Solarized + 부트스트랩 스타일)
$script:SettingsOpen = $false

function Set-RoundRegion {
    param($Control, [int]$Radius)
    $w = $Control.Width
    $h = $Control.Height
    if ($w -le $Radius * 2 -or $h -le $Radius * 2) { return }
    $d = $Radius * 2
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc(0, 0, $d, $d, 180, 90)
    $path.AddArc($w - $d - 1, 0, $d, $d, 270, 90)
    $path.AddArc($w - $d - 1, $h - $d - 1, $d, $d, 0, 90)
    $path.AddArc(0, $h - $d - 1, $d, $d, 90, 90)
    $path.CloseFigure()
    $Control.Region = New-Object System.Drawing.Region($path)
    $path.Dispose()
}

function New-BsButton {
    param([string]$Text, [bool]$Primary, [int]$X, [int]$Y, [int]$W, [int]$H)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.Location = New-Object System.Drawing.Point($X, $Y)
    $b.Size = New-Object System.Drawing.Size($W, $H)
    $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $b.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $b.UseVisualStyleBackColor = $false
    if ($Primary) {
        $b.BackColor = $script:UI.Blue
        $b.ForeColor = [System.Drawing.Color]::White
        $b.FlatAppearance.BorderColor = $script:UI.Blue
    } else {
        $b.BackColor = $script:UI.Base2
        $b.ForeColor = $script:UI.Base01
        $b.FlatAppearance.BorderColor = $script:UI.Base1
    }
    $b.FlatAppearance.BorderSize = 1
    Set-RoundRegion -Control $b -Radius (S 6)
    return $b
}

function New-ColorButton {
    param([System.Drawing.Color]$Color, [int]$X, [int]$Y)
    $b = New-Object System.Windows.Forms.Button
    $b.Location = New-Object System.Drawing.Point($X, $Y)
    $b.Size = New-Object System.Drawing.Size((S 64), (S 26))
    $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $b.FlatAppearance.BorderColor = $script:UI.Base1
    $b.BackColor = $Color
    $b.add_Click({
        param($s, $e)
        $cd = New-Object System.Windows.Forms.ColorDialog
        $cd.Color = $s.BackColor
        $cd.FullOpen = $true
        if ($cd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $s.BackColor = $cd.Color }
        $cd.Dispose()
    })
    Set-RoundRegion -Control $b -Radius (S 5)
    return $b
}

function Show-SettingsDialog {
    if ($script:SettingsOpen) { return }
    $script:SettingsOpen = $true
    try {
        $dlg = New-Object System.Windows.Forms.Form
        $dlg.Text = '블루투스 배터리 바 설정'
        $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $dlg.MaximizeBox = $false
        $dlg.MinimizeBox = $false
        $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $dlg.TopMost = $true
        $dlg.Font = New-Object System.Drawing.Font('Segoe UI', 9)
        $dlg.BackColor = $script:UI.Base3
        $dlg.ForeColor = $script:UI.Base01
        $dlg.ClientSize = New-Object System.Drawing.Size((S 420), (S 718))

        # --- 카드 1: 표시 방식 ---
        $gb1 = New-Object System.Windows.Forms.GroupBox
        $gb1.Text = '표시 방식'
        $gb1.ForeColor = $script:UI.Base02
        $gb1.Location = New-Object System.Drawing.Point((S 14), (S 12))
        $gb1.Size = New-Object System.Drawing.Size((S 392), (S 108))

        $rbNum = New-Object System.Windows.Forms.RadioButton
        $rbNum.Text = '숫자로 표시  (예: 85%)'
        $rbNum.ForeColor = $script:UI.Base01
        $rbNum.Location = New-Object System.Drawing.Point((S 14), (S 24))
        $rbNum.Size = New-Object System.Drawing.Size((S 360), (S 24))
        $rbNum.Checked = ($script:Settings.DisplayMode -eq 'number')

        $rbIcon = New-Object System.Windows.Forms.RadioButton
        $rbIcon.Text = '배터리 모양 아이콘으로 표시'
        $rbIcon.ForeColor = $script:UI.Base01
        $rbIcon.Location = New-Object System.Drawing.Point((S 14), (S 50))
        $rbIcon.Size = New-Object System.Drawing.Size((S 360), (S 24))
        $rbIcon.Checked = ($script:Settings.DisplayMode -eq 'icon')

        $rbIconNum = New-Object System.Windows.Forms.RadioButton
        $rbIconNum.Text = '배터리 아이콘 안에 숫자로 표시'
        $rbIconNum.ForeColor = $script:UI.Base01
        $rbIconNum.Location = New-Object System.Drawing.Point((S 14), (S 76))
        $rbIconNum.Size = New-Object System.Drawing.Size((S 360), (S 24))
        $rbIconNum.Checked = ($script:Settings.DisplayMode -eq 'iconNumber')

        $gb1.Controls.Add($rbNum)
        $gb1.Controls.Add($rbIcon)
        $gb1.Controls.Add($rbIconNum)
        $dlg.Controls.Add($gb1)

        # --- 카드 2: 색상 ---
        $gb2 = New-Object System.Windows.Forms.GroupBox
        $gb2.Text = '색상'
        $gb2.ForeColor = $script:UI.Base02
        $gb2.Location = New-Object System.Drawing.Point((S 14), (S 128))
        $gb2.Size = New-Object System.Drawing.Size((S 392), (S 202))

        $rbColor = New-Object System.Windows.Forms.RadioButton
        $rbColor.Text = '컬러  (20% 이하 경고색, 10% 이하 빨강)'
        $rbColor.ForeColor = $script:UI.Base01
        $rbColor.Location = New-Object System.Drawing.Point((S 14), (S 24))
        $rbColor.Size = New-Object System.Drawing.Size((S 340), (S 24))
        $rbColor.Checked = ($script:Settings.IconColorMode -ne 'gray')

        $rbGray = New-Object System.Windows.Forms.RadioButton
        $rbGray.Text = '그레이스케일  (테마 단색)'
        $rbGray.ForeColor = $script:UI.Base01
        $rbGray.Location = New-Object System.Drawing.Point((S 14), (S 50))
        $rbGray.Size = New-Object System.Drawing.Size((S 340), (S 24))
        $rbGray.Checked = ($script:Settings.IconColorMode -eq 'gray')

        $lbColor = New-Object System.Windows.Forms.Label
        $lbColor.Text = '기본 색상 (정상 잔량):'
        $lbColor.ForeColor = $script:UI.Base01
        $lbColor.AutoSize = $true
        $lbColor.Location = New-Object System.Drawing.Point((S 16), (S 84))

        $btnColor = New-ColorButton -Color (ConvertTo-ColorSafe $script:Settings.NormalColor ([System.Drawing.Color]::FromArgb(76, 175, 80))) -X (S 160) -Y (S 79)

        $lbWarn = New-Object System.Windows.Forms.Label
        $lbWarn.Text = '경고 색상 (20% 이하):'
        $lbWarn.ForeColor = $script:UI.Base01
        $lbWarn.AutoSize = $true
        $lbWarn.Location = New-Object System.Drawing.Point((S 16), (S 118))

        $btnWarn = New-ColorButton -Color (ConvertTo-ColorSafe $script:Settings.WarnColor ([System.Drawing.Color]::FromArgb(255, 205, 60))) -X (S 160) -Y (S 113)

        $lbText = New-Object System.Windows.Forms.Label
        $lbText.Text = '아이콘 숫자 색상:'
        $lbText.ForeColor = $script:UI.Base01
        $lbText.AutoSize = $true
        $lbText.Location = New-Object System.Drawing.Point((S 16), (S 152))

        $btnText = New-ColorButton -Color (Get-IconNumberColor) -X (S 160) -Y (S 147)

        $chkInvertText = New-Object System.Windows.Forms.CheckBox
        $chkInvertText.Text = '게이지 경계에서 숫자 색상 반전'
        $chkInvertText.ForeColor = $script:UI.Base01
        $chkInvertText.Location = New-Object System.Drawing.Point((S 238), (S 148))
        $chkInvertText.Size = New-Object System.Drawing.Size((S 142), (S 42))
        $chkInvertText.Checked = [bool]$script:Settings.InvertIconText

        $gb2.Controls.Add($rbColor)
        $gb2.Controls.Add($rbGray)
        $gb2.Controls.Add($lbColor)
        $gb2.Controls.Add($btnColor)
        $gb2.Controls.Add($lbWarn)
        $gb2.Controls.Add($btnWarn)
        $gb2.Controls.Add($lbText)
        $gb2.Controls.Add($btnText)
        $gb2.Controls.Add($chkInvertText)
        $dlg.Controls.Add($gb2)

        # --- 연결된 장치만 표시 ---
        $chkConn = New-Object System.Windows.Forms.CheckBox
        $chkConn.Text = '현재 연결된 장치만 표시'
        $chkConn.ForeColor = $script:UI.Base01
        $chkConn.Location = New-Object System.Drawing.Point((S 28), (S 336))
        $chkConn.Size = New-Object System.Drawing.Size((S 360), (S 24))
        $chkConn.Checked = [bool]$script:Settings.ConnectedOnly
        $dlg.Controls.Add($chkConn)

        # --- 카드 3: 장치 표시 설정 (표시 여부 + 표시 이름) ---
        $gb3 = New-Object System.Windows.Forms.GroupBox
        $gb3.Text = '장치 표시 설정  (표시 이름을 비우면 자동 약칭)'
        $gb3.ForeColor = $script:UI.Base02
        $gb3.Location = New-Object System.Drawing.Point((S 14), (S 366))
        $gb3.Size = New-Object System.Drawing.Size((S 392), (S 306))

        $grid = New-Object System.Windows.Forms.DataGridView
        $grid.Location = New-Object System.Drawing.Point((S 12), (S 24))
        $grid.Size = New-Object System.Drawing.Size((S 368), (S 268))
        $grid.AllowUserToAddRows = $false
        $grid.AllowUserToDeleteRows = $false
        $grid.AllowUserToResizeRows = $false
        $grid.RowHeadersVisible = $false
        $grid.MultiSelect = $false
        $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::CellSelect
        $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
        $grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
        $grid.ColumnHeadersHeight = (S 30)
        $grid.RowTemplate.Height = (S 28)
        $grid.BackgroundColor = $script:UI.Base3
        $grid.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $grid.GridColor = $script:UI.Base2
        $grid.EnableHeadersVisualStyles = $false
        $grid.ColumnHeadersDefaultCellStyle.BackColor = $script:UI.Base2
        $grid.ColumnHeadersDefaultCellStyle.ForeColor = $script:UI.Base02
        $grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = $script:UI.Base2
        $grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
        $grid.ColumnHeadersDefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
        $grid.ColumnHeadersDefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::False
        $grid.DefaultCellStyle.BackColor = $script:UI.Base3
        $grid.DefaultCellStyle.ForeColor = $script:UI.Base01
        $grid.DefaultCellStyle.SelectionBackColor = $script:UI.Sel
        $grid.DefaultCellStyle.SelectionForeColor = $script:UI.Base02
        $grid.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(4, 0, 4, 0)
        $grid.EditMode = [System.Windows.Forms.DataGridViewEditMode]::EditOnEnter

        $colShow = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
        $colShow.HeaderText = '표시'
        $colShow.Width = (S 58)
        $colShow.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
        $colShow.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $colName.HeaderText = '장치 이름'
        $colName.ReadOnly = $true
        $colName.FillWeight = 52
        $colAlias = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $colAlias.HeaderText = '표시 이름'
        $colAlias.FillWeight = 42
        [void]$grid.Columns.Add($colShow)
        [void]$grid.Columns.Add($colName)
        [void]$grid.Columns.Add($colAlias)
        $grid.Columns[0].DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter

        $grid.add_CurrentCellDirtyStateChanged({
            if ($grid.IsCurrentCellDirty) {
                [void]$grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
            }
        })
        $grid.add_CellPainting({
            param($s, $e)
            if ($e.RowIndex -lt 0 -or $e.ColumnIndex -ne 0) { return }
            $back = if (($e.State -band [System.Windows.Forms.DataGridViewElementStates]::Selected) -ne 0) { $script:UI.Sel } else { $script:UI.Base3 }
            $bg = New-Object System.Drawing.SolidBrush($back)
            $border = New-Object System.Drawing.Pen($script:UI.Base2)
            $e.Graphics.FillRectangle($bg, $e.CellBounds)
            $e.Graphics.DrawLine($border, $e.CellBounds.Right - 1, $e.CellBounds.Top, $e.CellBounds.Right - 1, $e.CellBounds.Bottom)
            $e.Graphics.DrawLine($border, $e.CellBounds.Left, $e.CellBounds.Bottom - 1, $e.CellBounds.Right, $e.CellBounds.Bottom - 1)

            $boxSize = [Math]::Min((S 16), $e.CellBounds.Height - (S 8))
            $boxX = $e.CellBounds.Left + [int](($e.CellBounds.Width - $boxSize) / 2)
            $boxY = $e.CellBounds.Top + [int](($e.CellBounds.Height - $boxSize) / 2)
            $boxRect = New-Object System.Drawing.Rectangle($boxX, $boxY, $boxSize, $boxSize)
            $checked = $false
            try { $checked = [System.Convert]::ToBoolean($e.Value) } catch { }

            $path = New-Object System.Drawing.Drawing2D.GraphicsPath
            $r = (S 4)
            $path.AddArc($boxRect.Left, $boxRect.Top, $r, $r, 180, 90)
            $path.AddArc($boxRect.Right - $r, $boxRect.Top, $r, $r, 270, 90)
            $path.AddArc($boxRect.Right - $r, $boxRect.Bottom - $r, $r, $r, 0, 90)
            $path.AddArc($boxRect.Left, $boxRect.Bottom - $r, $r, $r, 90, 90)
            $path.CloseFigure()
            if ($checked) {
                $fill = New-Object System.Drawing.SolidBrush($script:UI.Blue)
                $e.Graphics.FillPath($fill, $path)
                $fill.Dispose()
                $tick = New-Object System.Drawing.Pen([System.Drawing.Color]::White, (S 2))
                $tick.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
                $tick.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
                $p1 = New-Object System.Drawing.Point([int]($boxX + $boxSize * 0.25), [int]($boxY + $boxSize * 0.52))
                $p2 = New-Object System.Drawing.Point([int]($boxX + $boxSize * 0.43), [int]($boxY + $boxSize * 0.70))
                $p3 = New-Object System.Drawing.Point([int]($boxX + $boxSize * 0.76), [int]($boxY + $boxSize * 0.31))
                $e.Graphics.DrawLines($tick, @($p1, $p2, $p3))
                $tick.Dispose()
            } else {
                $empty = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
                $outline = New-Object System.Drawing.Pen($script:UI.Base1, 1)
                $e.Graphics.FillPath($empty, $path)
                $e.Graphics.DrawPath($outline, $path)
                $empty.Dispose(); $outline.Dispose()
            }
            $path.Dispose(); $border.Dispose(); $bg.Dispose()
            $e.Handled = $true
        })

        # 알려진 장치: 현재 감지 + 숨김 목록 + 별칭 보유 장치
        $names = New-Object 'System.Collections.Generic.List[string]'
        foreach ($item in $script:LastAll) { if (-not $names.Contains($item.Name)) { [void]$names.Add($item.Name) } }
        foreach ($h in @($script:Settings.HiddenDevices)) { if (-not [string]::IsNullOrWhiteSpace($h) -and -not $names.Contains($h)) { [void]$names.Add($h) } }
        foreach ($k in @($script:Settings.Aliases.Keys)) { if (-not $names.Contains($k)) { [void]$names.Add($k) } }
        $hidden = @($script:Settings.HiddenDevices)
        foreach ($n in ($names | Sort-Object)) {
            $alias = ''
            if ($script:Settings.Aliases.ContainsKey($n)) { $alias = [string]$script:Settings.Aliases[$n] }
            [void]$grid.Rows.Add(($hidden -notcontains $n), $n, $alias)
        }

        $gb3.Controls.Add($grid)
        $dlg.Controls.Add($gb3)

        # --- 버튼 ---
        $btnOk = New-BsButton -Text '확인' -Primary $true -X (S 236) -Y (S 684) -W (S 80) -H (S 32)
        $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $btnCancel = New-BsButton -Text '취소' -Primary $false -X (S 326) -Y (S 684) -W (S 80) -H (S 32)
        $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $dlg.Controls.Add($btnOk)
        $dlg.Controls.Add($btnCancel)
        $dlg.AcceptButton = $btnOk
        $dlg.CancelButton = $btnCancel

        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try { [void]$grid.EndEdit() } catch { }
            if ($rbIconNum.Checked) {
                $script:Settings.DisplayMode = 'iconNumber'
            } elseif ($rbIcon.Checked) {
                $script:Settings.DisplayMode = 'icon'
            } else {
                $script:Settings.DisplayMode = 'number'
            }
            $script:Settings.IconColorMode = if ($rbGray.Checked) { 'gray' } else { 'color' }
            $script:Settings.NormalColor = '#{0:X2}{1:X2}{2:X2}' -f $btnColor.BackColor.R, $btnColor.BackColor.G, $btnColor.BackColor.B
            $script:Settings.WarnColor = '#{0:X2}{1:X2}{2:X2}' -f $btnWarn.BackColor.R, $btnWarn.BackColor.G, $btnWarn.BackColor.B
            $script:Settings.IconTextColor = '#{0:X2}{1:X2}{2:X2}' -f $btnText.BackColor.R, $btnText.BackColor.G, $btnText.BackColor.B
            $script:Settings.InvertIconText = $chkInvertText.Checked
            $script:Settings.ConnectedOnly = $chkConn.Checked
            $newHidden = @()
            $newAliases = @{}
            foreach ($row in $grid.Rows) {
                $n = [string]$row.Cells[1].Value
                if ([string]::IsNullOrWhiteSpace($n)) { continue }
                if (-not [bool]$row.Cells[0].Value) { $newHidden += $n }
                $a = [string]$row.Cells[2].Value
                if (-not [string]::IsNullOrWhiteSpace($a)) { $newAliases[$n] = $a.Trim() }
            }
            $script:Settings.HiddenDevices = $newHidden
            $script:Settings.Aliases = $newAliases
            Save-Settings
            if ($null -ne $script:Form -and -not $script:Form.IsDisposed) { Update-Bar -Force }
        }
        $dlg.Dispose()
    } catch {
        Write-Log ('설정 창 오류: {0}' -f $_.Exception.Message)
    } finally {
        $script:SettingsOpen = $false
    }
}

# ------------------------------------------------------------ 메뉴
function Update-Menu {
    $script:Menu.Items.Clear()

    $header = $script:Menu.Items.Add('블루투스 배터리 바')
    $header.Enabled = $false
    [void]$script:Menu.Items.Add('-')

    $hidden = @($script:Settings.HiddenDevices)
    $shown = @($script:LastAll | Where-Object { $hidden -notcontains $_.Name -and ($_.Connected -or -not $script:Settings.ConnectedOnly) })
    if ($shown.Count -eq 0) {
        $none = $script:Menu.Items.Add('표시 중인 장치가 없습니다')
        $none.Enabled = $false
    } else {
        foreach ($item in $shown) {
            $mi = $script:Menu.Items.Add(('{0}   {1}' -f $item.Name, $item.Detail))
            $mi.Enabled = $false
        }
    }
    [void]$script:Menu.Items.Add('-')

    $refresh = $script:Menu.Items.Add('지금 새로고침')
    $refresh.add_Click({ Update-Bar -Force })

    $period = New-Object System.Windows.Forms.ToolStripMenuItem('새로고침 주기')
    foreach ($sec in 30, 60, 120, 300) {
        $label = if ($sec -lt 60) { '{0}초' -f $sec } else { '{0}분' -f ($sec / 60) }
        $sub = New-Object System.Windows.Forms.ToolStripMenuItem($label)
        $sub.Tag = $sec
        $sub.Checked = ($sec -eq $script:Settings.RefreshSec)
        $sub.add_Click({
            param($s, $e)
            $script:Settings.RefreshSec = [int]$s.Tag
            $script:Timer.Interval = $script:Settings.RefreshSec * 1000
            Save-Settings
            Update-Menu
        })
        [void]$period.DropDownItems.Add($sub)
    }
    [void]$script:Menu.Items.Add($period)

    $cfg = $script:Menu.Items.Add('설정...')
    $cfg.add_Click({ Show-SettingsDialog })

    $reset = $script:Menu.Items.Add('위치 초기화 (오른쪽 끝)')
    $reset.add_Click({
        $script:Settings.BarOffsetRight = 0
        Save-Settings
        Position-Bar
    })

    $auto = New-Object System.Windows.Forms.ToolStripMenuItem('Windows 시작 시 자동 실행')
    $auto.Checked = (Test-Path $script:StartupLnk)
    $auto.add_Click({
        param($s, $e)
        if (Test-Path $script:StartupLnk) {
            Remove-Item $script:StartupLnk -Force -ErrorAction SilentlyContinue
        } else {
            try {
                $ws = New-Object -ComObject WScript.Shell
                $sc = $ws.CreateShortcut($script:StartupLnk)
                $sc.TargetPath = 'wscript.exe'
                $sc.Arguments = '"{0}"' -f (Join-Path $script:BaseDir 'start_bar.vbs')
                $sc.WorkingDirectory = $script:BaseDir
                $sc.Description = '블루투스 배터리 바'
                $sc.Save()
            } catch {
                Write-Log ('자동 실행 등록 실패: {0}' -f $_.Exception.Message)
            }
        }
        Update-Menu
    })
    [void]$script:Menu.Items.Add($auto)
    [void]$script:Menu.Items.Add('-')

    $quit = $script:Menu.Items.Add('종료')
    $quit.add_Click({ Stop-Bar })
}

function Stop-Bar {
    try { $script:Timer.Stop() } catch { }
    try { $script:WatchTimer.Stop() } catch { }
    try { $script:OnceTimer.Stop() } catch { }
    try { $script:Form.Visible = $false } catch { }
    [System.Windows.Forms.Application]::Exit()
}

# ------------------------------------------------------------ 폼 생성
function New-BarForm {
    $f = New-Object System.Windows.Forms.Form
    $f.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $f.ShowInTaskbar = $false
    $f.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
    $f.Size = New-Object System.Drawing.Size(120, 34)
    $f.Location = New-Object System.Drawing.Point(-3000, -3000)   # 화면 밖에서 시작 (깜빡임 방지)
    $f.Text = 'BtBatteryBar'
    $f.TopMost = $true
    $f.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 28)
    Add-DragHandlers -Control $f
    return $f
}

# ------------------------------------------------------------ 테스트 모드
if ($Test) {
    Write-Output '=== 블루투스 배터리 조회 테스트 (바 버전) ==='
    Read-Settings
    Write-Output ('설정: ConnectedOnly={0}, DisplayMode={1}, IconColorMode={2}, NormalColor={3}, WarnColor={4}, IconTextColor={5}, InvertIconText={6}' -f `
        $script:Settings.ConnectedOnly, $script:Settings.DisplayMode, $script:Settings.IconColorMode, $script:Settings.NormalColor, $script:Settings.WarnColor, $script:Settings.IconTextColor, $script:Settings.InvertIconText)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $list = Get-BtBatteryList
    $sw.Stop()
    Write-Output ('조회 시간: {0} ms / 장치 {1}개 (Connected=True 만 바에 표시됨)' -f $sw.ElapsedMilliseconds, $list.Count)
    $list | Format-Table Short, Percent, Connected, Detail, Name -AutoSize | Out-String | Write-Output

    $h = [BarApi]::FindWindow('Shell_TrayWnd', [NullString]::Value)
    Write-Output ('Shell_TrayWnd 핸들: {0}' -f $h)
    if ($h -ne [IntPtr]::Zero) {
        $rect = New-Object 'BarApi+RECT'
        [void][BarApi]::GetWindowRect($h, [ref]$rect)
        Write-Output ('작업표시줄 영역: {0},{1} ~ {2},{3} (높이 {4}px)' -f $rect.Left, $rect.Top, $rect.Right, $rect.Bottom, ($rect.Bottom - $rect.Top))
        $notify = [BarApi]::FindWindowEx($h, [IntPtr]::Zero, 'TrayNotifyWnd', [NullString]::Value)
        if ($notify -ne [IntPtr]::Zero) {
            $nr = New-Object 'BarApi+RECT'
            [void][BarApi]::GetWindowRect($notify, [ref]$nr)
            Write-Output ('알림 영역: {0},{1} ~ {2},{3}  →  우측 기준점 X = {4}' -f $nr.Left, $nr.Top, $nr.Right, $nr.Bottom, ($nr.Left - 10))
        }
    }

    Write-Output ''
    Write-Output '=== 배터리 아이콘 렌더링 테스트 (3배 확대 PNG, 현재 색상 설정 기준) ==='
    $outDir = Join-Path $env:TEMP 'BtBatteryBarTest'
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
    $palDark = Get-Palette -Light $false
    $palLight = Get-Palette -Light $true
    foreach ($p in 8, 15, 35, 75, 100) {
        foreach ($mode in 'dark', 'light') {
            $pal = if ($mode -eq 'dark') { $palDark } else { $palLight }
            foreach ($style in 'plain', 'number') {
                $showNumber = ($style -eq 'number')
                $imgW = if ($showNumber) { 44 } else { 36 }
                $imgH = if ($showNumber) { 24 } else { 22 }
                $img = New-BatteryImage -Percent $p -Pal $pal -W $imgW -H $imgH -ShowPercent:$showNumber
                $big = New-Object System.Drawing.Bitmap(($imgW * 3), ($imgH * 3))
                $gb = [System.Drawing.Graphics]::FromImage($big)
                if ($mode -eq 'dark') { $gb.Clear([System.Drawing.Color]::FromArgb(38, 38, 38)) } else { $gb.Clear([System.Drawing.Color]::FromArgb(232, 232, 232)) }
                $gb.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::NearestNeighbor
                $gb.DrawImage($img, 0, 0, ($imgW * 3), ($imgH * 3))
                $gb.Dispose()
                $path = Join-Path $outDir ('battery_{0}_{1}_{2}.png' -f $p, $mode, $style)
                $big.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
                $big.Dispose(); $img.Dispose()
                Write-Output $path
            }
        }
    }
    exit 0
}

# ------------------------------------------------------------ 설정창 단독 테스트
if ($SettingsTest) {
    Read-Settings
    $script:LastAll = Get-BtBatteryList
    $script:Menu = New-Object System.Windows.Forms.ContextMenuStrip   # Add-DragHandlers 참조용
    Show-SettingsDialog
    Write-Output ('저장된 설정: DisplayMode={0}, IconColorMode={1}, NormalColor={2}, WarnColor={3}, IconTextColor={4}, InvertIconText={5}, ConnectedOnly={6}' -f `
        $script:Settings.DisplayMode, $script:Settings.IconColorMode, $script:Settings.NormalColor, $script:Settings.WarnColor, $script:Settings.IconTextColor, $script:Settings.InvertIconText, $script:Settings.ConnectedOnly)
    Write-Output ('Hidden=[{0}] / Aliases: {1}' -f (@($script:Settings.HiddenDevices) -join ', '), (($script:Settings.Aliases.Keys | ForEach-Object { '{0}->{1}' -f $_, $script:Settings.Aliases[$_] }) -join ', '))
    exit 0
}

# ------------------------------------------------------------ 메인
$script:Mutex = New-Object System.Threading.Mutex($false, 'BtBatteryBar_SingleInstance')
if (-not $script:Mutex.WaitOne(0, $false)) {
    [void][System.Windows.Forms.MessageBox]::Show('블루투스 배터리 바가 이미 실행 중입니다.', '블루투스 배터리 바')
    exit 0
}

try {
    Read-Settings

    $script:Tip = New-Object System.Windows.Forms.ToolTip
    $script:Menu = New-Object System.Windows.Forms.ContextMenuStrip
    $script:Form = New-BarForm
    Update-Menu

    # 핵심 순서: top-level 상태로 먼저 표시한 뒤 위치 지정
    # (TopLevel=false + SetParent 방식은 Win11 25H2 에서 차단되어 사용하지 않음)
    $script:Form.Show()
    Set-BarExStyle
    Update-Bar -Force

    # 배터리 갱신 타이머
    $script:Timer = New-Object System.Windows.Forms.Timer
    $script:Timer.Interval = $script:Settings.RefreshSec * 1000
    $script:Timer.add_Tick({ Update-Bar })
    $script:Timer.Start()

    # 더블클릭 새로고침용 지연 원샷 타이머
    $script:OnceTimer = New-Object System.Windows.Forms.Timer
    $script:OnceTimer.Interval = 150
    $script:OnceTimer.add_Tick({
        $script:OnceTimer.Stop()
        Update-Bar -Force
    })

    # 감시 타이머 (작업표시줄 추적 / 전체화면 숨김 / 최상위 유지)
    $script:WatchTimer = New-Object System.Windows.Forms.Timer
    $script:WatchTimer.Interval = 1000
    $script:WatchTimer.add_Tick({ Watch-Tick })
    $script:WatchTimer.Start()

    $ctx = New-Object System.Windows.Forms.ApplicationContext
    [System.Windows.Forms.Application]::Run($ctx)
} catch {
    Write-Log ('치명적 오류(바): {0}' -f $_.Exception.Message)
    Write-Log $_.ScriptStackTrace
} finally {
    try { $script:Form.Dispose() } catch { }
    try { $script:Mutex.ReleaseMutex() } catch { }
    $script:Mutex.Dispose()
}
