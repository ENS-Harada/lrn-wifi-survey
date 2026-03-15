# ============================================================
#  ライフリズムナビ Wi-Fi電波測定ツール
#  営業が各居室を回って電波強度を測定するためのツール
# ============================================================
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# ----- 定数 -----
$MEASURE_COUNT = 10       # 1居室あたりの測定回数
$MEASURE_INTERVAL = 2     # 測定間隔（秒）

# ----- 電波強度の評価関数 -----
function Get-SignalRating {
    param([int]$signal)
    if ($signal -ge 80) { return "優良" }
    elseif ($signal -ge 60) { return "良好" }
    elseif ($signal -ge 40) { return "普通" }
    elseif ($signal -ge 20) { return "弱い" }
    else { return "非常に弱い" }
}

function Get-SignalColor {
    param([string]$rating)
    switch ($rating) {
        "優良"     { return "#10b981" }
        "良好"     { return "#3b82f6" }
        "普通"     { return "#f59e0b" }
        "弱い"     { return "#ef4444" }
        "非常に弱い" { return "#991b1b" }
        default    { return "#94a3b8" }
    }
}

# ----- 周辺SSID一覧を取得する関数 -----
function Get-NearbyNetworks {
    $networks = @()
    try {
        $output = netsh wlan show networks mode=bssid 2>$null
        $current = @{ SSID = ""; BSSID = ""; Signal = 0; Channel = 0; Band = ""; RadioType = ""; Auth = "" }
        $inEntry = $false

        foreach ($line in $output) {
            if ($line -match '^SSID\s+\d+\s+:\s*(.*)$') {
                $current = @{ SSID = $Matches[1].Trim(); BSSID = ""; Signal = 0; Channel = 0; Band = ""; RadioType = ""; Auth = "" }
                $inEntry = $true
            }
            if ($inEntry) {
                if ($line -match '^\s+BSSID\s+\d+\s+:\s+(.+)$') {
                    # 前のBSSIDエントリがあれば保存
                    if ($current.BSSID -ne "") {
                        $networks += [PSCustomObject]$current
                        $current = @{ SSID = $current.SSID; BSSID = ""; Signal = 0; Channel = 0; Band = ""; RadioType = ""; Auth = $current.Auth }
                    }
                    $current.BSSID = $Matches[1].Trim()
                }
                if ($line -match '(?:シグナル|Signal)\s+:\s+(\d+)%') { $current.Signal = [int]$Matches[1] }
                if ($line -match '(?:無線の種類|Radio type)\s+:\s+(.+)$') { $current.RadioType = $Matches[1].Trim() }
                if ($line -match '(?:チャネル|Channel)\s+:\s+(\d+)') {
                    $ch = [int]$Matches[1]
                    $current.Channel = $ch
                    $current.Band = if ($ch -le 14) { "2.4GHz" } else { "5GHz" }
                }
                if ($line -match '(?:認証|Authentication)\s+:\s+(.+)$') { $current.Auth = $Matches[1].Trim() }
            }
        }
        # 最後のエントリ
        if ($current.BSSID -ne "") { $networks += [PSCustomObject]$current }
    } catch { }

    return $networks
}

# ----- チャネル干渉分析関数 -----
function Get-ChannelAnalysis {
    param($networks)

    $analysis = @{ Channels24 = @{}; Channels5 = @{}; Recommendation24 = ""; Recommendation5 = "" }

    # 2.4GHz チャネル（1-14）
    for ($ch = 1; $ch -le 14; $ch++) { $analysis.Channels24[$ch] = @{ Count = 0; MaxSignal = 0; Networks = @() } }
    # 5GHz 主要チャネル
    foreach ($ch in @(36,40,44,48,52,56,60,64,100,104,108,112,116,120,124,128,132,136,140,144,149,153,157,161,165)) {
        $analysis.Channels5[$ch] = @{ Count = 0; MaxSignal = 0; Networks = @() }
    }

    foreach ($nw in $networks) {
        $ch = $nw.Channel
        if ($nw.Band -eq "2.4GHz" -and $analysis.Channels24.ContainsKey($ch)) {
            $analysis.Channels24[$ch].Count++
            if ($nw.Signal -gt $analysis.Channels24[$ch].MaxSignal) { $analysis.Channels24[$ch].MaxSignal = $nw.Signal }
            $analysis.Channels24[$ch].Networks += "$($nw.SSID)($($nw.Signal)%)"
        }
        elseif ($nw.Band -eq "5GHz" -and $analysis.Channels5.ContainsKey($ch)) {
            $analysis.Channels5[$ch].Count++
            if ($nw.Signal -gt $analysis.Channels5[$ch].MaxSignal) { $analysis.Channels5[$ch].MaxSignal = $nw.Signal }
            $analysis.Channels5[$ch].Networks += "$($nw.SSID)($($nw.Signal)%)"
        }
    }

    # 2.4GHz推奨チャネル（1, 6, 11のみ判定。重なりが少ないものを推奨）
    $bestCh24 = @(1, 6, 11) | Sort-Object { $analysis.Channels24[$_].Count }, { $analysis.Channels24[$_].MaxSignal } | Select-Object -First 1
    $analysis.Recommendation24 = "ch$bestCh24（AP数: $($analysis.Channels24[$bestCh24].Count)）"

    # 5GHz推奨チャネル（W52: 36,40,44,48 を優先）
    $usedCh5 = $analysis.Channels5.GetEnumerator() | Where-Object { $_.Value.Count -gt 0 } | Sort-Object { $_.Value.Count } | Select-Object -First 1
    $emptyCh5 = $analysis.Channels5.GetEnumerator() | Where-Object { $_.Value.Count -eq 0 -and $_.Key -le 48 } | Select-Object -First 1
    if ($emptyCh5) {
        $analysis.Recommendation5 = "ch$($emptyCh5.Key)（空き - W52帯域）"
    } elseif ($usedCh5) {
        $analysis.Recommendation5 = "ch$($usedCh5.Key)（利用が最も少ない）"
    } else {
        $analysis.Recommendation5 = "データなし"
    }

    return $analysis
}

# ----- Wi-Fi情報を1回取得する関数 -----
function Get-WiFiInfo {
    $info = @{
        SSID      = ""
        Signal    = 0
        Channel   = ""
        RadioType = ""
        Auth      = ""
        Cipher    = ""
        BSSID     = ""
        Band      = ""
    }

    try {
        $output = netsh wlan show interfaces 2>$null
        if ($output) {
            foreach ($line in $output) {
                if ($line -match '^\s+SSID\s+:\s+(.+)$') { $info.SSID = $Matches[1].Trim() }
                if ($line -match '^\s+BSSID\s+:\s+(.+)$') { $info.BSSID = $Matches[1].Trim() }
                if ($line -match '(?:シグナル|Signal)\s+:\s+(\d+)%') { $info.Signal = [int]$Matches[1] }
                if ($line -match '(?:チャネル|Channel)\s+:\s+(\d+)') {
                    if ($line -notmatch '(?:無線|Radio)') { $info.Channel = $Matches[1].Trim() }
                }
                if ($line -match '(?:無線の種類|Radio type)\s+:\s+(.+)$') { $info.RadioType = $Matches[1].Trim() }
                if ($line -match '(?:認証|Authentication)\s+:\s+(.+)$') { $info.Auth = $Matches[1].Trim() }
                if ($line -match '(?:暗号|Cipher)\s+:\s+(.+)$') { $info.Cipher = $Matches[1].Trim() }
            }
            # チャネル番号から周波数帯を判定
            if ($info.Channel -ne "") {
                $ch = [int]$info.Channel
                $info.Band = if ($ch -le 14) { "2.4GHz" } else { "5GHz" }
            }
        }
    } catch { }

    return $info
}

# ============================================================
#  メイン処理
# ============================================================
Clear-Host
Write-Host ""
Write-Host "  ╔══════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║  ライフリズムナビ Wi-Fi電波測定ツール    ║" -ForegroundColor Cyan
Write-Host "  ║  居室を回りながら電波強度を測定します    ║" -ForegroundColor Cyan
Write-Host "  ╚══════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# Wi-Fi接続確認
$checkWifi = Get-WiFiInfo
if ($checkWifi.SSID -eq "") {
    Write-Host "  ⚠ Wi-Fiに接続されていません。" -ForegroundColor Red
    Write-Host "    Wi-Fiに接続してから再度実行してください。" -ForegroundColor Red
    Write-Host ""
    Read-Host "  Enterキーで終了"
    exit
}

Write-Host "  接続中のWi-Fi: $($checkWifi.SSID)" -ForegroundColor Green
Write-Host "  周波数帯: $($checkWifi.Band) / チャネル: $($checkWifi.Channel)" -ForegroundColor Green
Write-Host ""

# 施設名の入力
$facilityName = ""
while ($facilityName -eq "") {
    $facilityName = Read-Host "  施設名を入力してください"
    if ($facilityName -eq "") { Write-Host "  ※ 施設名は必須です" -ForegroundColor Yellow }
}

Write-Host ""
Write-Host "  周辺Wi-Fiネットワークをスキャン中..." -ForegroundColor Cyan
$nearbyNetworks = Get-NearbyNetworks
$channelAnalysis = Get-ChannelAnalysis $nearbyNetworks

$nw24 = $nearbyNetworks | Where-Object { $_.Band -eq "2.4GHz" }
$nw5  = $nearbyNetworks | Where-Object { $_.Band -eq "5GHz" }

Write-Host "  検出: $($nearbyNetworks.Count) AP（2.4GHz: $($nw24.Count) / 5GHz: $($nw5.Count)）" -ForegroundColor Green
Write-Host "  2.4GHz推奨: $($channelAnalysis.Recommendation24)" -ForegroundColor White
Write-Host "  5GHz推奨:   $($channelAnalysis.Recommendation5)" -ForegroundColor White
Write-Host ""
Write-Host "  ────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  各居室に移動して、居室名を入力してください。" -ForegroundColor White
Write-Host "  測定が終わったら「q」と入力して終了します。" -ForegroundColor White
Write-Host "  ────────────────────────────────────" -ForegroundColor DarkGray
Write-Host ""

# 測定データの配列
$allResults = @()
$roomIndex = 0

# ----- 居室ごとの測定ループ -----
while ($true) {
    $roomName = Read-Host "  居室名を入力（終了: q）"
    if ($roomName -eq "q" -or $roomName -eq "Q") { break }
    if ($roomName -eq "") {
        Write-Host "  ※ 居室名を入力してください" -ForegroundColor Yellow
        continue
    }

    $roomIndex++
    Write-Host ""
    Write-Host "  [$roomIndex] 「$roomName」を測定中..." -ForegroundColor Cyan

    $measurements = @()

    for ($i = 1; $i -le $MEASURE_COUNT; $i++) {
        $wifi = Get-WiFiInfo
        $measurements += $wifi.Signal

        # プログレスバー表示
        $bar = "█" * $i + "░" * ($MEASURE_COUNT - $i)
        $rating = Get-SignalRating $wifi.Signal
        Write-Host "`r  [$bar] $i/$MEASURE_COUNT  信号: $($wifi.Signal)%（$rating）  " -NoNewline

        if ($i -lt $MEASURE_COUNT) { Start-Sleep -Seconds $MEASURE_INTERVAL }
    }

    # 統計計算
    $avg = [math]::Round(($measurements | Measure-Object -Average).Average, 1)
    $min = ($measurements | Measure-Object -Minimum).Minimum
    $max = ($measurements | Measure-Object -Maximum).Maximum
    $avgRating = Get-SignalRating ([int]$avg)

    Write-Host ""
    Write-Host ""
    Write-Host "  ┌─────────────────────────────────┐" -ForegroundColor Green
    Write-Host "  │  結果: 平均 $avg%（$avgRating）" -ForegroundColor Green
    Write-Host "  │  最小: ${min}%  最大: ${max}%" -ForegroundColor Green
    Write-Host "  └─────────────────────────────────┘" -ForegroundColor Green
    Write-Host ""

    # 結果を保存
    $wifiInfo = Get-WiFiInfo
    $allResults += [PSCustomObject]@{
        No         = $roomIndex
        RoomName   = $roomName
        SSID       = $wifiInfo.SSID
        Band       = $wifiInfo.Band
        Channel    = $wifiInfo.Channel
        RadioType  = $wifiInfo.RadioType
        AvgSignal  = $avg
        MinSignal  = $min
        MaxSignal  = $max
        Rating     = $avgRating
        Auth       = $wifiInfo.Auth
        Cipher     = $wifiInfo.Cipher
        BSSID      = $wifiInfo.BSSID
        Timestamp  = Get-Date -Format "HH:mm:ss"
        RawData    = ($measurements -join ",")
    }
}

# ----- 結果がない場合 -----
if ($allResults.Count -eq 0) {
    Write-Host ""
    Write-Host "  測定データがありません。終了します。" -ForegroundColor Yellow
    Read-Host "  Enterキーで終了"
    exit
}

# ============================================================
#  出力ファイル生成
# ============================================================
Write-Host ""
Write-Host "  レポートを生成中..." -ForegroundColor Cyan

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptDir) { $scriptDir = Get-Location }
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$dateStr = Get-Date -Format "yyyy/MM/dd HH:mm"
$csvFile = Join-Path $scriptDir "WiFi_Survey_${timestamp}.csv"
$htmlFile = Join-Path $scriptDir "WiFi_Survey_${timestamp}.html"

# ----- CSV出力 -----
$csvHeader = "No,居室名,SSID,周波数帯,チャネル,無線規格,平均信号強度(%),最小(%),最大(%),評価,認証方式,暗号化,BSSID,測定時刻,測定生データ"
$csvLines = @($csvHeader)
foreach ($r in $allResults) {
    $csvLines += "$($r.No),`"$($r.RoomName)`",`"$($r.SSID)`",$($r.Band),$($r.Channel),`"$($r.RadioType)`",$($r.AvgSignal),$($r.MinSignal),$($r.MaxSignal),`"$($r.Rating)`",`"$($r.Auth)`",`"$($r.Cipher)`",`"$($r.BSSID)`",$($r.Timestamp),`"$($r.RawData)`""
}
[System.IO.File]::WriteAllLines($csvFile, $csvLines, [System.Text.Encoding]::GetEncoding("shift_jis"))

# ----- 全体サマリー計算 -----
$overallAvg = [math]::Round(($allResults | ForEach-Object { $_.AvgSignal } | Measure-Object -Average).Average, 1)
$overallMin = ($allResults | ForEach-Object { $_.MinSignal } | Measure-Object -Minimum).Minimum
$overallRating = Get-SignalRating ([int]$overallAvg)
$weakRooms = ($allResults | Where-Object { $_.AvgSignal -lt 40 }).Count

# ----- HTMLレポート生成 -----
$tableRows = ""
foreach ($r in $allResults) {
    $color = Get-SignalColor $r.Rating
    $barWidth = [math]::Min($r.AvgSignal, 100)
    $tableRows += @"
<tr>
<td style="text-align:center">$($r.No)</td>
<td><strong>$($r.RoomName)</strong></td>
<td>$($r.SSID)</td>
<td style="text-align:center">$($r.Band)</td>
<td style="text-align:center">$($r.Channel)</td>
<td>
  <div style="display:flex;align-items:center;gap:8px">
    <div style="width:120px;height:20px;background:#e5e7eb;border-radius:10px;overflow:hidden">
      <div style="width:${barWidth}%;height:100%;background:${color};border-radius:10px"></div>
    </div>
    <span style="font-weight:700;color:${color}">$($r.AvgSignal)%</span>
  </div>
</td>
<td style="text-align:center">$($r.MinSignal)%</td>
<td style="text-align:center">$($r.MaxSignal)%</td>
<td style="text-align:center"><span style="color:${color};font-weight:700">$($r.Rating)</span></td>
<td style="text-align:center;font-size:12px;color:#94a3b8">$($r.Timestamp)</td>
</tr>
"@
}

$overallColor = Get-SignalColor $overallRating
$weakWarning = ""
if ($weakRooms -gt 0) {
    $weakWarning = "<div style='background:#fef2f2;border:2px solid #ef4444;border-radius:8px;padding:12px 16px;margin-top:16px;color:#991b1b;font-weight:600'>&#9888; 電波が弱い居室が ${weakRooms} 箇所あります。AP（アクセスポイント）の追加を検討してください。</div>"
}

# ----- 周辺ネットワーク一覧HTML生成 -----
$nearbyTableRows = ""
$nearbyIndex = 0
foreach ($nw in ($nearbyNetworks | Sort-Object Band, { -($_.Signal) })) {
    $nearbyIndex++
    $sigColor = Get-SignalColor (Get-SignalRating $nw.Signal)
    $barW = [math]::Min($nw.Signal, 100)
    $nearbyTableRows += @"
<tr>
<td style="text-align:center;font-size:12px">$nearbyIndex</td>
<td><strong>$(if($nw.SSID){"$($nw.SSID)"}else{"（非公開）"})</strong></td>
<td style="font-size:11px;font-family:monospace">$($nw.BSSID)</td>
<td style="text-align:center"><span style="background:$(if($nw.Band -eq '5GHz'){'#dbeafe;color:#2563eb'}else{'#fef3c7;color:#d97706'});padding:2px 8px;border-radius:4px;font-size:11px;font-weight:600">$($nw.Band)</span></td>
<td style="text-align:center">$($nw.Channel)</td>
<td>
  <div style="display:flex;align-items:center;gap:6px">
    <div style="width:80px;height:14px;background:#e5e7eb;border-radius:7px;overflow:hidden">
      <div style="width:${barW}%;height:100%;background:${sigColor};border-radius:7px"></div>
    </div>
    <span style="font-size:12px;font-weight:600;color:${sigColor}">$($nw.Signal)%</span>
  </div>
</td>
<td style="font-size:11px">$($nw.Auth)</td>
</tr>
"@
}

# ----- チャネル混雑度HTML（2.4GHz）-----
$ch24Html = ""
foreach ($ch in @(1,2,3,4,5,6,7,8,9,10,11,12,13)) {
    $info = $channelAnalysis.Channels24[$ch]
    $cnt = $info.Count
    $bgColor = if ($cnt -eq 0) { "#d1fae5" } elseif ($cnt -le 2) { "#fef3c7" } elseif ($cnt -le 5) { "#fed7aa" } else { "#fecaca" }
    $fgColor = if ($cnt -eq 0) { "#059669" } elseif ($cnt -le 2) { "#d97706" } elseif ($cnt -le 5) { "#ea580c" } else { "#dc2626" }
    $highlight = if ($ch -eq 1 -or $ch -eq 6 -or $ch -eq 11) { "font-weight:700;border:2px solid $fgColor;" } else { "" }
    $ch24Html += "<div style='text-align:center;padding:8px 4px;background:$bgColor;border-radius:6px;min-width:45px;$highlight' title='$($info.Networks -join ", ")'><div style='font-size:10px;color:#64748b'>ch$ch</div><div style='font-size:18px;font-weight:700;color:$fgColor'>$cnt</div><div style='font-size:9px;color:#94a3b8'>AP</div></div>"
}

# ----- チャネル混雑度HTML（5GHz）-----
$ch5Html = ""
foreach ($ch in @(36,40,44,48,52,56,60,64,100,104,108,112,116,120,124,128,132,136,140,149,153,157,161,165)) {
    $info = $channelAnalysis.Channels5[$ch]
    if (-not $info) { continue }
    $cnt = $info.Count
    $bgColor = if ($cnt -eq 0) { "#d1fae5" } elseif ($cnt -le 1) { "#fef3c7" } else { "#fecaca" }
    $fgColor = if ($cnt -eq 0) { "#059669" } elseif ($cnt -le 1) { "#d97706" } else { "#dc2626" }
    $bandLabel = if ($ch -le 48) { "W52" } elseif ($ch -le 64) { "W53" } elseif ($ch -le 144) { "W56" } else { "W52ext" }
    $ch5Html += "<div style='text-align:center;padding:6px 3px;background:$bgColor;border-radius:6px;min-width:40px' title='${bandLabel}: $($info.Networks -join ", ")'><div style='font-size:9px;color:#64748b'>$ch</div><div style='font-size:14px;font-weight:700;color:$fgColor'>$cnt</div></div>"
}

$html = @"
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Wi-Fi電波測定レポート - $facilityName</title>
<style>
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family:'Segoe UI','Meiryo','Yu Gothic',sans-serif; background:#f0f4f8; color:#1a202c; padding:20px; }
.container { max-width:1100px; margin:0 auto; }
.header { background:linear-gradient(135deg,#7c3aed,#5b21b6); color:#fff; padding:24px 32px; border-radius:12px; margin-bottom:20px; }
.header h1 { font-size:22px; font-weight:700; }
.header p { font-size:13px; opacity:0.85; margin-top:4px; }
.summary { display:flex; gap:16px; margin-bottom:20px; flex-wrap:wrap; }
.summary-card { flex:1; min-width:180px; background:#fff; border-radius:12px; padding:20px; box-shadow:0 2px 8px rgba(0,0,0,0.08); text-align:center; }
.summary-card .label { font-size:12px; color:#64748b; font-weight:600; text-transform:uppercase; }
.summary-card .value { font-size:32px; font-weight:700; margin-top:4px; }
.summary-card .sub { font-size:13px; color:#94a3b8; margin-top:2px; }
.card { background:#fff; border-radius:12px; box-shadow:0 2px 8px rgba(0,0,0,0.08); margin-bottom:20px; overflow:hidden; }
.card-header { padding:16px 24px; border-bottom:1px solid #e5e7eb; }
.card-header h2 { font-size:16px; font-weight:700; color:#1e293b; }
table { width:100%; border-collapse:collapse; }
th { text-align:left; padding:10px 12px; background:#f8fafc; color:#64748b; font-size:11px; font-weight:600; text-transform:uppercase; letter-spacing:0.5px; border-bottom:2px solid #e2e8f0; }
td { padding:10px 12px; font-size:13px; border-bottom:1px solid #f1f5f9; }
.legend { display:flex; gap:20px; padding:16px 24px; flex-wrap:wrap; border-top:1px solid #e5e7eb; }
.legend-item { display:flex; align-items:center; gap:6px; font-size:12px; color:#64748b; }
.legend-dot { width:12px; height:12px; border-radius:50%; }
.footer { text-align:center; padding:16px; color:#94a3b8; font-size:12px; }
.note { font-size:12px; color:#64748b; padding:12px 24px; background:#f8fafc; border-top:1px solid #e5e7eb; }
.btn-group { display:flex; gap:12px; justify-content:center; padding:24px; flex-wrap:wrap; }
.btn { padding:12px 32px; border-radius:8px; font-size:14px; font-weight:600; cursor:pointer; border:none; transition:all 0.2s; }
.btn-primary { background:#7c3aed; color:#fff; }
.btn-primary:hover { background:#6d28d9; }
.btn-secondary { background:#f1f5f9; color:#475569; }
.btn-secondary:hover { background:#e2e8f0; }
.channel-grid { display:flex; gap:4px; flex-wrap:wrap; }
@media print { body{background:#fff;padding:10px;} .btn-group{display:none;} .summary-card{box-shadow:none;border:1px solid #e5e7eb;} .card{box-shadow:none;border:1px solid #e5e7eb;} .channel-grid{gap:2px;} }
</style>
</head>
<body>
<div class="container">

<div class="header">
<h1>&#128225; Wi-Fi電波測定レポート</h1>
<p>施設名: $facilityName ／ 測定日時: $dateStr ／ 測定居室数: $($allResults.Count)</p>
</div>

<div class="summary">
<div class="summary-card">
<div class="label">全体平均</div>
<div class="value" style="color:$overallColor">${overallAvg}%</div>
<div class="sub">$overallRating</div>
</div>
<div class="summary-card">
<div class="label">最低値</div>
<div class="value" style="color:#ef4444">${overallMin}%</div>
<div class="sub">全居室の最小値</div>
</div>
<div class="summary-card">
<div class="label">測定居室数</div>
<div class="value" style="color:#2563eb">$($allResults.Count)</div>
<div class="sub">各${MEASURE_COUNT}回測定</div>
</div>
<div class="summary-card">
<div class="label">電波弱 居室数</div>
<div class="value" style="color:$(if($weakRooms -gt 0){'#ef4444'}else{'#10b981'})">$weakRooms</div>
<div class="sub">$(if($weakRooms -gt 0){'AP追加を検討'}else{'問題なし'})</div>
</div>
</div>

$weakWarning

<div class="card">
<div class="card-header"><h2>&#128246; 居室別 測定結果</h2></div>
<table>
<tr>
<th style="width:40px">No</th>
<th>居室名</th>
<th>SSID</th>
<th style="text-align:center">帯域</th>
<th style="text-align:center">Ch</th>
<th>平均信号強度</th>
<th style="text-align:center">最小</th>
<th style="text-align:center">最大</th>
<th style="text-align:center">評価</th>
<th style="text-align:center">時刻</th>
</tr>
$tableRows
</table>
<div class="legend">
<div class="legend-item"><div class="legend-dot" style="background:#10b981"></div> 優良（80%以上）</div>
<div class="legend-item"><div class="legend-dot" style="background:#3b82f6"></div> 良好（60-79%）</div>
<div class="legend-item"><div class="legend-dot" style="background:#f59e0b"></div> 普通（40-59%）</div>
<div class="legend-item"><div class="legend-dot" style="background:#ef4444"></div> 弱い（20-39%）</div>
<div class="legend-item"><div class="legend-dot" style="background:#991b1b"></div> 非常に弱い（20%未満）</div>
</div>
<div class="note">
※ 信号強度はWindows標準のnetshコマンドで取得（0-100%）。各居室${MEASURE_COUNT}回測定の統計値です。<br>
※ ライフリズムナビは5GHz帯を推奨しています。2.4GHz帯の場合は5GHz対応APへの変更をご検討ください。
</div>
</div>

<div class="card">
<div class="card-header"><h2>&#128225; 周辺ネットワーク一覧</h2></div>
<table>
<tr>
<th style="width:30px">No</th>
<th>SSID</th>
<th>BSSID</th>
<th style="text-align:center">帯域</th>
<th style="text-align:center">Ch</th>
<th>信号強度</th>
<th>認証方式</th>
</tr>
$nearbyTableRows
</table>
<div class="note">
※ スキャン実行時に検出された周辺のWi-Fiアクセスポイント一覧です。信号強度が強い他APは電波干渉の原因になります。
</div>
</div>

<div class="card">
<div class="card-header"><h2>&#128202; チャネル干渉分析</h2></div>
<div style="padding:20px 24px">
<h3 style="font-size:14px;font-weight:700;color:#1e293b;margin-bottom:12px">2.4GHz帯 チャネル利用状況</h3>
<div class="channel-grid">
$ch24Html
</div>
<div style="display:flex;gap:16px;flex-wrap:wrap;margin:10px 0 4px;font-size:11px;color:#64748b">
<span style="display:flex;align-items:center;gap:4px"><span style="width:10px;height:10px;border-radius:50%;background:#d1fae5;display:inline-block"></span> 空き（0）</span>
<span style="display:flex;align-items:center;gap:4px"><span style="width:10px;height:10px;border-radius:50%;background:#fef3c7;display:inline-block"></span> 少ない（1-2 AP）</span>
<span style="display:flex;align-items:center;gap:4px"><span style="width:10px;height:10px;border-radius:50%;background:#fed7aa;display:inline-block"></span> やや多い（3-5 AP）</span>
<span style="display:flex;align-items:center;gap:4px"><span style="width:10px;height:10px;border-radius:50%;background:#fecaca;display:inline-block"></span> 利用が多い（6+ AP）</span>
</div>
<div style="display:flex;align-items:center;gap:8px;margin:8px 0 4px;font-size:12px;color:#64748b">
<span>&#9733; 推奨チャネル（利用が少ない）: </span>
<span style="font-weight:700;color:#059669">$($channelAnalysis.Recommendation24)</span>
</div>
<div style="font-size:11px;color:#94a3b8;margin-bottom:24px">※ ch1 / ch6 / ch11 が非重複チャネル（太枠表示）。同一チャネルのAP数が多いほど干渉リスクが高まります。</div>

<h3 style="font-size:14px;font-weight:700;color:#1e293b;margin-bottom:12px">5GHz帯 チャネル利用状況</h3>
<div class="channel-grid">
$ch5Html
</div>
<div style="display:flex;gap:16px;flex-wrap:wrap;margin:10px 0 4px;font-size:11px;color:#64748b">
<span style="display:flex;align-items:center;gap:4px"><span style="width:10px;height:10px;border-radius:50%;background:#d1fae5;display:inline-block"></span> 空き（0）</span>
<span style="display:flex;align-items:center;gap:4px"><span style="width:10px;height:10px;border-radius:50%;background:#fef3c7;display:inline-block"></span> 利用あり（1 AP）</span>
<span style="display:flex;align-items:center;gap:4px"><span style="width:10px;height:10px;border-radius:50%;background:#fecaca;display:inline-block"></span> 利用が多い（2+ AP）</span>
</div>
<div style="display:flex;align-items:center;gap:8px;margin:8px 0 4px;font-size:12px;color:#64748b">
<span>&#9733; 推奨チャネル: </span>
<span style="font-weight:700;color:#059669">$($channelAnalysis.Recommendation5)</span>
</div>
<div style="font-size:11px;color:#94a3b8;margin-bottom:8px">※ W52(36-48) / W53(52-64) / W56(100-140)。5GHz帯はチャネル数が多いため、1APでも同一チャネルに乗っていれば注意が必要です。</div>
</div>
<div class="note">
※ 各チャネルの数字は検出されたAP数です。AP数が多く信号が強いチャネルほど干渉の影響を受けやすくなります。<br>
※ ライフリズムナビは5GHz帯を推奨しています。5GHz帯で空きチャネル（W52推奨）の利用を検討してください。
</div>
</div>

<div class="btn-group">
<button class="btn btn-primary" onclick="window.print()">&#128424; 印刷 / PDF保存</button>
<button class="btn btn-secondary" onclick="location.reload()">&#128260; 閉じる</button>
</div>

<div class="footer">
ライフリズムナビ Wi-Fi電波測定ツール v1.1 &mdash; エコナビスタ株式会社
</div>

</div>
</body>
</html>
"@

[System.IO.File]::WriteAllText($htmlFile, $html, [System.Text.Encoding]::UTF8)

# ----- 完了表示 -----
Write-Host ""
Write-Host "  ╔══════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "  ║  測定完了！レポートを生成しました        ║" -ForegroundColor Green
Write-Host "  ╚══════════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""
Write-Host "  Excel(CSV): $csvFile" -ForegroundColor White
Write-Host "  HTML:       $htmlFile" -ForegroundColor White
Write-Host ""

# ブラウザで開く
Start-Process $htmlFile

Write-Host "  ブラウザでレポートが開きます。" -ForegroundColor Gray
Write-Host ""
Read-Host "  Enterキーで終了"
