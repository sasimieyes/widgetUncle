# ============================================================
# BtBatteryWidget.ps1
# 블루투스 장치별 배터리 잔량을 작업표시줄 트레이에 표시하는 위젯
# - 장치마다 트레이 아이콘 1개 (숫자 + 하단 게이지 바)
# - 우클릭 메뉴: 장치 목록 / 새로고침 / 주기 변경 / 자동 시작 / 종료
# - 사용법: wscript start_widget.vbs (창 없이 실행)
#   테스트: powershell -STA -File BtBatteryWidget.ps1 -Test
# ============================================================
param(
    [switch]$Test
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -MemberDefinition '[DllImport("user32.dll", SetLastError = true)] public static extern bool DestroyIcon(IntPtr hIcon);' -Name 'IconUtil' -Namespace 'Win32'

$script:BaseDir      = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:SettingsPath = Join-Path $script:BaseDir 'settings.json'
$script:LogPath      = Join-Path $script:BaseDir 'widget.log'
$script:StartupLnk   = Join-Path ([Environment]::GetFolderPath('Startup')) 'BtBatteryWidget.lnk'
$script:BatteryKey   = '{104EA319-6EE2-4701-BD47-8DDBF425BBE5} 2'  # DEVPKEY_Bluetooth_Battery

# ------------------------------------------------------------ 설정
$script:Settings = @{ RefreshSec = 60; AlertPercent = 15 }

function Read-Settings {
    if (-not (Test-Path $script:SettingsPath)) { return }
    try {
        $j = Get-Content $script:SettingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($j.RefreshSec -ge 10)   { $script:Settings.RefreshSec   = [int]$j.RefreshSec }
        if ($j.AlertPercent -ge 1)  { $script:Settings.AlertPercent = [int]$j.AlertPercent }
    } catch { }
}

function Save-Settings {
    try {
        $obj = [PSCustomObject]@{
            RefreshSec   = $script:Settings.RefreshSec
            AlertPercent = $script:Settings.AlertPercent
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
# 반환: @( @{ Name=장치명; Percent=대표값(최소); Detail='85% / 87%' }, ... )
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
        try {
            $resp = Invoke-CimMethod -InputObject $dev -MethodName GetDeviceProperties `
                -Arguments @{ devicePropertyKeys = @($script:BatteryKey) } -ErrorAction Stop
            $prop = @($resp.deviceProperties) | Where-Object { $null -ne $_.Data } | Select-Object -First 1
            if ($prop) { $pct = [int]$prop.Data }
        } catch { continue }
        if ($null -eq $pct -or $pct -lt 0 -or $pct -gt 100) { continue }
        if (-not $byName.ContainsKey($dev.Name)) {
            $byName[$dev.Name] = New-Object System.Collections.ArrayList
        }
        [void]$byName[$dev.Name].Add($pct)
    }
    $list = foreach ($name in ($byName.Keys | Sort-Object)) {
        $vals = @($byName[$name] | Sort-Object)
        $uniq = @($vals | Select-Object -Unique)
        [PSCustomObject]@{
            Name    = $name
            Percent = $vals[0]
            Detail  = (($uniq | ForEach-Object { '{0}%' -f $_ }) -join ' / ')
        }
    }
    return @($list)
}

# ------------------------------------------------------------ 아이콘 그리기
function Get-TaskbarIsLight {
    try {
        $v = Get-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize' `
            -Name SystemUsesLightTheme -ErrorAction Stop
        return ($v.SystemUsesLightTheme -eq 1)
    } catch { return $false }
}

# 32x32 비트맵: 위쪽에 % 숫자, 아래쪽에 잔량 게이지 바
function New-BatteryBitmap {
    param([int]$Percent, [bool]$LightTaskbar)
    $bmp = New-Object System.Drawing.Bitmap(32, 32, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb))
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAlias
        $g.Clear([System.Drawing.Color]::Transparent)

        if ($Percent -le $script:Settings.AlertPercent) {
            $mainColor = [System.Drawing.Color]::FromArgb(255, 82, 82)     # 빨강
        } elseif ($Percent -le 40) {
            $mainColor = [System.Drawing.Color]::FromArgb(255, 170, 60)    # 주황
        } elseif ($LightTaskbar) {
            $mainColor = [System.Drawing.Color]::FromArgb(25, 25, 25)      # 라이트 테마용 검정
        } else {
            $mainColor = [System.Drawing.Color]::White                     # 다크 테마용 흰색
        }

        $fontSize = if ($Percent -ge 100) { 13 } else { 18 }
        $font  = New-Object System.Drawing.Font('Segoe UI', $fontSize, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
        $brush = New-Object System.Drawing.SolidBrush($mainColor)
        $fmt   = New-Object System.Drawing.StringFormat
        $fmt.Alignment     = [System.Drawing.StringAlignment]::Center
        $fmt.LineAlignment = [System.Drawing.StringAlignment]::Center
        $rect = New-Object System.Drawing.RectangleF(0, -1, 32, 27)
        $g.DrawString([string]$Percent, $font, $brush, $rect, $fmt)

        $track = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(90, 128, 128, 128))
        $g.FillRectangle($track, 2, 27, 28, 4)
        $fillW = [Math]::Max(1, [int][Math]::Round(28.0 * $Percent / 100))
        $fill = New-Object System.Drawing.SolidBrush($mainColor)
        $g.FillRectangle($fill, 2, 27, $fillW, 4)

        $font.Dispose(); $brush.Dispose(); $fmt.Dispose(); $track.Dispose(); $fill.Dispose()
        return $bmp
    } finally {
        $g.Dispose()
    }
}

function New-BatteryIconHandle {
    param([int]$Percent, [bool]$LightTaskbar)
    $bmp = New-BatteryBitmap -Percent $Percent -LightTaskbar $LightTaskbar
    try {
        return $bmp.GetHicon()
    } finally {
        $bmp.Dispose()
    }
}

# ------------------------------------------------------------ 트레이 아이콘 관리
# $script:Tray : 장치명 -> @{ Notify; Hicon; IconObj; Alerted }
$script:Tray = @{}
$script:LastList = @()

function Set-EntryIcon {
    param($Entry, [int]$Percent, [bool]$LightTaskbar)
    $newHandle = New-BatteryIconHandle -Percent $Percent -LightTaskbar $LightTaskbar
    $newIcon = [System.Drawing.Icon]::FromHandle($newHandle)
    $Entry.Notify.Icon = $newIcon
    if ($Entry.Hicon -ne [IntPtr]::Zero) { [void][Win32.IconUtil]::DestroyIcon($Entry.Hicon) }
    if ($null -ne $Entry.IconObj) { $Entry.IconObj.Dispose() }
    $Entry.Hicon = $newHandle
    $Entry.IconObj = $newIcon
}

function Remove-Entry {
    param([string]$Name)
    $e = $script:Tray[$Name]
    if ($null -eq $e) { return }
    $e.Notify.Visible = $false
    $e.Notify.Dispose()
    if ($e.Hicon -ne [IntPtr]::Zero) { [void][Win32.IconUtil]::DestroyIcon($e.Hicon) }
    if ($null -ne $e.IconObj) { $e.IconObj.Dispose() }
    $script:Tray.Remove($Name)
}

function Update-Widget {
    param([switch]$Force)
    if (-not $Force -and $script:Menu.Visible) { return }   # 메뉴 사용 중에는 다음 주기로 미룸

    $light = Get-TaskbarIsLight
    $list = Get-BtBatteryList
    $script:LastList = $list
    $names = @($list | ForEach-Object { $_.Name })

    foreach ($key in @($script:Tray.Keys)) {
        if ($names -notcontains $key) { Remove-Entry -Name $key }
    }

    foreach ($item in $list) {
        $tip = '{0}  {1}' -f $item.Name, $item.Detail
        if ($tip.Length -gt 63) { $tip = $tip.Substring(0, 60) + '...' }

        if (-not $script:Tray.ContainsKey($item.Name)) {
            $ni = New-Object System.Windows.Forms.NotifyIcon
            $ni.ContextMenuStrip = $script:Menu
            $ni.add_DoubleClick({ Update-Widget -Force })
            $entry = @{ Notify = $ni; Hicon = [IntPtr]::Zero; IconObj = $null; Alerted = $false }
            $script:Tray[$item.Name] = $entry
            Set-EntryIcon -Entry $entry -Percent $item.Percent -LightTaskbar $light
            $ni.Text = $tip
            $ni.Visible = $true
        } else {
            $entry = $script:Tray[$item.Name]
            Set-EntryIcon -Entry $entry -Percent $item.Percent -LightTaskbar $light
            $entry.Notify.Text = $tip
        }

        # 저전량 풍선 알림 (임계값 이하로 떨어질 때 1회)
        $entry = $script:Tray[$item.Name]
        if ($item.Percent -le $script:Settings.AlertPercent) {
            if (-not $entry.Alerted) {
                $entry.Alerted = $true
                $msg = '{0} 배터리가 {1}% 남았습니다. 충전해 주세요.' -f $item.Name, $item.Percent
                $entry.Notify.ShowBalloonTip(5000, '배터리 부족', $msg, [System.Windows.Forms.ToolTipIcon]::Warning)
            }
        } elseif ($item.Percent -ge ($script:Settings.AlertPercent + 10)) {
            $entry.Alerted = $false
        }
    }

    Update-Menu
}

# ------------------------------------------------------------ 컨텍스트 메뉴
function Update-Menu {
    $script:Menu.Items.Clear()

    $header = $script:Menu.Items.Add('블루투스 배터리 위젯')
    $header.Enabled = $false
    [void]$script:Menu.Items.Add('-')

    if ($script:LastList.Count -eq 0) {
        $none = $script:Menu.Items.Add('배터리 정보를 제공하는 장치가 없습니다')
        $none.Enabled = $false
    } else {
        foreach ($item in $script:LastList) {
            $mi = $script:Menu.Items.Add(('{0}   {1}' -f $item.Name, $item.Detail))
            $mi.Enabled = $false
        }
    }
    [void]$script:Menu.Items.Add('-')

    $refresh = $script:Menu.Items.Add('지금 새로고침')
    $refresh.add_Click({ Update-Widget -Force })

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
                $sc.Arguments = '"{0}"' -f (Join-Path $script:BaseDir 'start_widget.vbs')
                $sc.WorkingDirectory = $script:BaseDir
                $sc.Description = '블루투스 배터리 위젯'
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
    $quit.add_Click({ Stop-Widget })
}

function Stop-Widget {
    try { $script:Timer.Stop() } catch { }
    foreach ($key in @($script:Tray.Keys)) { Remove-Entry -Name $key }
    [System.Windows.Forms.Application]::Exit()
}

# ------------------------------------------------------------ 테스트 모드
if ($Test) {
    Write-Output '=== 블루투스 배터리 조회 테스트 ==='
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $list = Get-BtBatteryList
    $sw.Stop()
    Write-Output ('조회 시간: {0} ms / 장치 {1}개' -f $sw.ElapsedMilliseconds, $list.Count)
    $list | Format-Table Name, Percent, Detail -AutoSize | Out-String | Write-Output

    Write-Output '=== 아이콘 렌더링 테스트 (4배 확대 PNG) ==='
    $outDir = Join-Path $env:TEMP 'BtBatteryWidgetTest'
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
    foreach ($p in 8, 35, 75, 100) {
        $small = New-BatteryBitmap -Percent $p -LightTaskbar $false
        $big = New-Object System.Drawing.Bitmap(128, 128)
        $gb = [System.Drawing.Graphics]::FromImage($big)
        $gb.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::NearestNeighbor
        $gb.DrawImage($small, 0, 0, 128, 128)
        $gb.Dispose()
        $path = Join-Path $outDir ('icon_{0}.png' -f $p)
        $big.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
        $big.Dispose(); $small.Dispose()
        Write-Output $path
    }
    exit 0
}

# ------------------------------------------------------------ 메인
$script:Mutex = New-Object System.Threading.Mutex($false, 'BtBatteryWidget_SingleInstance')
if (-not $script:Mutex.WaitOne(0, $false)) {
    [void][System.Windows.Forms.MessageBox]::Show('블루투스 배터리 위젯이 이미 실행 중입니다.', '블루투스 배터리 위젯')
    exit 0
}

try {
    Read-Settings

    $script:Menu = New-Object System.Windows.Forms.ContextMenuStrip
    Update-Menu

    $script:Timer = New-Object System.Windows.Forms.Timer
    $script:Timer.Interval = $script:Settings.RefreshSec * 1000
    $script:Timer.add_Tick({ Update-Widget })

    Update-Widget -Force
    $script:Timer.Start()

    $ctx = New-Object System.Windows.Forms.ApplicationContext
    [System.Windows.Forms.Application]::Run($ctx)
} catch {
    Write-Log ('치명적 오류: {0}' -f $_.Exception.Message)
    Write-Log $_.ScriptStackTrace
} finally {
    foreach ($key in @($script:Tray.Keys)) {
        try { Remove-Entry -Name $key } catch { }
    }
    try { $script:Mutex.ReleaseMutex() } catch { }
    $script:Mutex.Dispose()
}
