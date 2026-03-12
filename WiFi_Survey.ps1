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
@media print { body{background:#fff;padding:10px;} .btn-group{display:none;} .summary-card{box-shadow:none;border:1px solid #e5e7eb;} .card{box-shadow:none;border:1px solid #e5e7eb;} }
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

<div class="btn-group">
<button class="btn btn-primary" onclick="window.print()">&#128424; 印刷 / PDF保存</button>
<button class="btn btn-secondary" onclick="location.reload()">&#128260; 閉じる</button>
</div>

<div class="footer">
ライフリズムナビ Wi-Fi電波測定ツール v1.0 &mdash; エコナビスタ株式会社
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
