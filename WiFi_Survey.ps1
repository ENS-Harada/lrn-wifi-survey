# ============================================================
#  ライフリズムナビ Wi-Fi電波測定ツール v2.0
#  営業が各居室を回って電波強度を測定するためのツール
#
#  v2.0 変更点:
#    - 信号強度を RSSI(dBm) 表示に変更（Windows % 表示を廃止）
#    - 測定方式をスキャンベースに変更（ローミング影響を排除）
#    - 「推奨」表現を削除
#    - チャネル利用状況（AP台数）を居室ごとに表示（NW担当者向け参考情報）
#    - 周辺ネットワーク情報を居室ごとに取得・表示
# ============================================================
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# ----- 定数 -----
$MEASURE_COUNT    = 10   # 1居室あたりのスキャン回数
$MEASURE_INTERVAL = 2    # スキャン間隔（秒）

# ----- Windows signal(%) → RSSI(dBm) 変換 -----
function Convert-ToRSSI {
    param([int]$pct)
    # Windows signal% は dBm = (pct/2) - 100 で近似変換
    return [int](($pct / 2.0) - 100)
}

# ----- RSSI(dBm) 評価関数 -----
function Get-SignalRating {
    param([int]$rssi)
    if    ($rssi -ge -65) { return "優良" }
    elseif($rssi -ge -70) { return "良好" }
    elseif($rssi -ge -80) { return "普通" }
    elseif($rssi -ge -90) { return "弱い" }
    else                  { return "非常に弱い" }
}

function Get-SignalColor {
    param([string]$rating)
    switch ($rating) {
        "優良"       { return "#10b981" }
        "良好"       { return "#3b82f6" }
        "普通"       { return "#f59e0b" }
        "弱い"       { return "#ef4444" }
        "非常に弱い" { return "#991b1b" }
        default      { return "#94a3b8" }
    }
}

# ----- dBm 値から横バー表示幅を計算（0-100%） -----
# -100 dBm → 0%、-50 dBm → 100%
function Get-BarWidth {
    param([int]$rssi)
    $w = [int](($rssi + 100) * 2)
    return [math]::Max(0, [math]::Min(100, $w))
}

# ----- 周辺SSID一覧をスキャンする関数 -----
function Get-NearbyNetworks {
    $networks = @()
    try {
        $output = netsh wlan show networks mode=bssid 2>$null
        $current = @{ SSID = ""; BSSID = ""; SignalPct = 0; RSSI = -100; Channel = 0; Band = ""; RadioType = ""; Auth = "" }
        $inEntry = $false

        foreach ($line in $output) {
            if ($line -match '^SSID\s+\d+\s+:\s*(.*)$') {
                $current = @{ SSID = $Matches[1].Trim(); BSSID = ""; SignalPct = 0; RSSI = -100; Channel = 0; Band = ""; RadioType = ""; Auth = "" }
                $inEntry = $true
            }
            if ($inEntry) {
                if ($line -match '^\s+BSSID\s+\d+\s+:\s+(.+)$') {
                    if ($current.BSSID -ne "") {
                        $networks += [PSCustomObject]$current
                        $current = @{ SSID = $current.SSID; BSSID = ""; SignalPct = 0; RSSI = -100; Channel = 0; Band = ""; RadioType = ""; Auth = $current.Auth }
                    }
                    $current.BSSID = $Matches[1].Trim()
                }
                if ($line -match '(?:シグナル|Signal)\s+:\s+(\d+)%') {
                    $pct = [int]$Matches[1]
                    $current.SignalPct = $pct
                    $current.RSSI = Convert-ToRSSI $pct
                }
                if ($line -match '(?:無線の種類|Radio type)\s+:\s+(.+)$') { $current.RadioType = $Matches[1].Trim() }
                if ($line -match '(?:チャネル|Channel)\s+:\s+(\d+)') {
                    $ch = [int]$Matches[1]
                    $current.Channel = $ch
                    $current.Band = if ($ch -le 14) { "2.4GHz" } else { "5GHz" }
                }
                if ($line -match '(?:認証|Authentication)\s+:\s+(.+)$') { $current.Auth = $Matches[1].Trim() }
            }
        }
        if ($current.BSSID -ne "") { $networks += [PSCustomObject]$current }
    } catch { }

    return $networks
}

# ----- 接続中のSSIDを取得（初期値の参考表示用） -----
function Get-ConnectedSSID {
    try {
        $output = netsh wlan show interfaces 2>$null
        foreach ($line in $output) {
            if ($line -match '^\s+SSID\s+:\s+(.+)$') { return $Matches[1].Trim() }
        }
    } catch { }
    return ""
}

# ----- 接続中Wi-Fiの詳細情報を取得（スキャン未検出時のフォールバック用） -----
function Get-WiFiInfo {
    $info = @{ SSID = ""; RSSI = -100; Channel = 0; Band = ""; RadioType = ""; BSSID = "" }
    try {
        $output = netsh wlan show interfaces 2>$null
        foreach ($line in $output) {
            if ($line -match '^\s+SSID\s+:\s+(.+)$')                { $info.SSID      = $Matches[1].Trim() }
            if ($line -match '^\s+BSSID\s+:\s+(.+)$')               { $info.BSSID     = $Matches[1].Trim() }
            if ($line -match '(?:シグナル|Signal)\s+:\s+(\d+)%')    { $info.RSSI      = Convert-ToRSSI ([int]$Matches[1]) }
            if ($line -match '(?:無線の種類|Radio type)\s+:\s+(.+)$') { $info.RadioType = $Matches[1].Trim() }
            if ($line -match '(?:チャネル|Channel)\s+:\s+(\d+)') {
                if ($line -notmatch '(?:無線|Radio)') {
                    $ch = [int]$Matches[1]
                    $info.Channel = $ch
                    $info.Band = if ($ch -le 14) { "2.4GHz" } else { "5GHz" }
                }
            }
        }
    } catch { }
    return $info
}

# ----- チャネル利用状況分析 -----
function Get-ChannelAnalysis {
    param($networks)

    $analysis = @{ Channels24 = @{}; Channels5 = @{} }
    for ($ch = 1; $ch -le 14; $ch++) {
        $analysis.Channels24[$ch] = @{ Count = 0; MaxRSSI = -100; Networks = @() }
    }
    foreach ($ch in @(36,40,44,48,52,56,60,64,100,104,108,112,116,120,124,128,132,136,140,144,149,153,157,161,165)) {
        $analysis.Channels5[$ch] = @{ Count = 0; MaxRSSI = -100; Networks = @() }
    }

    foreach ($nw in $networks) {
        $ch = $nw.Channel
        if ($nw.Band -eq "2.4GHz" -and $analysis.Channels24.ContainsKey($ch)) {
            $analysis.Channels24[$ch].Count++
            if ($nw.RSSI -gt $analysis.Channels24[$ch].MaxRSSI) { $analysis.Channels24[$ch].MaxRSSI = $nw.RSSI }
            $analysis.Channels24[$ch].Networks += "$($nw.SSID)($($nw.RSSI)dBm)"
        }
        elseif ($nw.Band -eq "5GHz" -and $analysis.Channels5.ContainsKey($ch)) {
            $analysis.Channels5[$ch].Count++
            if ($nw.RSSI -gt $analysis.Channels5[$ch].MaxRSSI) { $analysis.Channels5[$ch].MaxRSSI = $nw.RSSI }
            $analysis.Channels5[$ch].Networks += "$($nw.SSID)($($nw.RSSI)dBm)"
        }
    }

    return $analysis
}

# ----- 干渉レベル評価 -----
# 自社SSID以外の同チャネル・隣接チャネルAPの数と信号強度で判定
function Get-InterferenceLevel {
    param([int]$targetChannel, [string]$targetBand, $nearbyNetworks, [string]$targetSSID)

    if ($targetChannel -eq 0 -or $targetBand -eq "") { return "不明" }

    $interferenceCount = 0
    $strongCount = 0

    foreach ($nw in $nearbyNetworks) {
        if ($nw.SSID -eq $targetSSID) { continue }   # 自社APは除外
        if ($nw.Band -ne $targetBand)  { continue }

        $chDiff = [math]::Abs($nw.Channel - $targetChannel)
        $isInterfering = $false

        if ($targetBand -eq "2.4GHz") {
            if ($chDiff -le 4) { $isInterfering = $true }   # 重複チャネル範囲（±4ch以内）
        } else {
            if ($chDiff -eq 0) { $isInterfering = $true }   # 5GHz は同一チャネルのみ
        }

        if ($isInterfering) {
            $interferenceCount++
            if ($nw.RSSI -ge -75) { $strongCount++ }  # -75dBm以上は強干渉源
        }
    }

    if    ($strongCount -ge 2 -or $interferenceCount -ge 4) { return "高" }
    elseif($strongCount -ge 1 -or $interferenceCount -ge 2) { return "中" }
    elseif($interferenceCount -ge 1)                        { return "低" }
    else                                                     { return "なし" }
}

function Get-InterferenceColor {
    param([string]$level)
    switch ($level) {
        "高"   { return "#ef4444" }
        "中"   { return "#f59e0b" }
        "低"   { return "#3b82f6" }
        "なし" { return "#10b981" }
        default { return "#94a3b8" }
    }
}

# ----- 居室ごとの営業向け一言コメント生成 -----
# 判定は電波強度（-70 dBm 社内目安値）のみに基づく
# チャネル干渉については業界共通の閾値が存在しないため当社では評価しない
function Get-RoomComment {
    param([string]$rating)

    # 信号強度コメント（測定値の事実のみ記述・-70 dBm が社内目安値）
    $sigComment = switch ($rating) {
        "優良"       { "ベッドフレーム付近・扉閉鎖の状態で、社内目安値（-70 dBm）を大きく上回る電波強度が計測されました。" }
        "良好"       { "ベッドフレーム付近・扉閉鎖の状態で、社内目安値（-70 dBm）を上回る電波強度が計測されました。" }
        "普通"       { "ベッドフレーム付近・扉閉鎖の状態で、社内目安値（-70 dBm）を下回る電波強度が計測されました。NW担当者への確認事項としてお伝えください。" }
        "弱い"       { "ベッドフレーム付近・扉閉鎖の状態で、社内目安値（-70 dBm）を大きく下回る電波強度が計測されました。NW担当者への確認事項としてお伝えください。" }
        "非常に弱い" { "ベッドフレーム付近・扉閉鎖の状態で、電波がほとんど届かない値が計測されました。NW担当者への確認事項としてお伝えください。" }
        default      { "" }
    }

    # 総合判定（-70 dBm 社内目安値のみを基準とする）
    $verdict = "確認事項なし"
    $verdictColor = "#10b981"
    $verdictBg    = "#f0fdf4"
    $verdictBorder = "#86efac"

    if ($rating -eq "非常に弱い" -or $rating -eq "弱い" -or $rating -eq "普通") {
        # -70 dBm 未達 → 社内基準未達
        $verdict = "確認事項あり"; $verdictColor = "#dc2626"; $verdictBg = "#fef2f2"; $verdictBorder = "#fca5a5"
    }

    return @{
        Signal        = $sigComment
        Verdict       = $verdict
        VerdictColor  = $verdictColor
        VerdictBg     = $verdictBg
        VerdictBorder = $verdictBorder
    }
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

# 初回スキャン
Write-Host "  周辺Wi-Fiネットワークをスキャン中..." -ForegroundColor Cyan
$initialScan = Get-NearbyNetworks

if ($initialScan.Count -eq 0) {
    Write-Host "  ⚠ Wi-Fiネットワークが検出されません。" -ForegroundColor Red
    Write-Host "    Wi-Fiアダプタが有効か確認してください。" -ForegroundColor Red
    Write-Host ""
    Read-Host "  Enterキーで終了"
    exit
}

# 接続中SSIDを取得
$connectedSSID = Get-ConnectedSSID

# SSID一覧を作成（スキャン結果 ＋ 接続中SSIDが未検出なら追加）
$ssidList = @($initialScan | Select-Object -ExpandProperty SSID -Unique | Where-Object { $_ -ne "" } | Sort-Object)
if ($connectedSSID -ne "" -and $ssidList -notcontains $connectedSSID) {
    $ssidList = @($connectedSSID) + $ssidList   # 接続中SSIDを先頭に追加
}

# 番号付きで一覧表示
Write-Host ""
Write-Host "  検出されたWi-Fiネットワーク:" -ForegroundColor White
$idx = 0
$defaultIdx = 0
foreach ($ssid in $ssidList) {
    $idx++
    $bestAP = $initialScan | Where-Object { $_.SSID -eq $ssid } | Sort-Object RSSI -Descending | Select-Object -First 1
    $rssiStr = if ($bestAP) { "ch:$($bestAP.Channel) $($bestAP.Band) $($bestAP.RSSI) dBm" } else { "（接続中・スキャン未検出）" }
    $marker  = ""
    if ($ssid -eq $connectedSSID) { $marker = " ◀ 接続中"; $defaultIdx = $idx }
    Write-Host "  [$idx] $ssid  $rssiStr$marker" -ForegroundColor $(if($ssid -eq $connectedSSID){"Green"}else{"Gray"})
}
Write-Host ""

# 番号で選択（Enterだけで接続中SSIDを選択）
$targetSSID = ""
$promptText = if ($defaultIdx -gt 0) { "  番号を入力してください（Enter で [$defaultIdx] を選択）" } else { "  番号を入力してください" }
while ($targetSSID -eq "") {
    $inputNum = Read-Host $promptText
    $inputNum = $inputNum.Trim()
    # Enterのみの場合は接続中SSIDをデフォルト選択
    if ($inputNum -eq "" -and $defaultIdx -gt 0) {
        $targetSSID = $ssidList[$defaultIdx - 1]
    } elseif ($inputNum -match '^\d+$') {
        $n = [int]$inputNum
        if ($n -ge 1 -and $n -le $ssidList.Count) {
            $targetSSID = $ssidList[$n - 1]
        } else {
            Write-Host "  ※ 1 〜 $($ssidList.Count) の番号を入力してください" -ForegroundColor Yellow
        }
    } else {
        Write-Host "  ※ 番号を入力してください（例: 1）" -ForegroundColor Yellow
    }
}
Write-Host "  → 測定対象: $targetSSID" -ForegroundColor Cyan

# 施設名入力
$facilityName = ""
while ($facilityName -eq "") {
    $facilityName = Read-Host "  施設名を入力してください"
    if ($facilityName -eq "") { Write-Host "  ※ 施設名は必須です" -ForegroundColor Yellow }
}

Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Yellow
Write-Host "  ║  【測定前の準備】各居室に入る前にご確認ください      ║" -ForegroundColor Yellow
Write-Host "  ╠══════════════════════════════════════════════════════╣" -ForegroundColor Yellow
Write-Host "  ║  ① 居室の扉を閉めた状態で測定してください           ║" -ForegroundColor Yellow
Write-Host "  ║  ② PCを介護ベッドのフレーム付近（床近く）に          ║" -ForegroundColor Yellow
Write-Host "  ║     置いた状態で測定してください                     ║" -ForegroundColor Yellow
Write-Host "  ║     （スリープセンサの実設置位置を想定）             ║" -ForegroundColor Yellow
Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Yellow
Write-Host ""
Write-Host "  測定が終わったら「q」を入力して終了します。" -ForegroundColor White
Write-Host ""

$allResults = @()
$roomIndex  = 0

# ----- 居室ごとの測定ループ -----
while ($true) {
    Write-Host "  ┌─────────────────────────────────────────┐" -ForegroundColor DarkYellow
    Write-Host "  │ 準備確認: 扉を閉めましたか？            │" -ForegroundColor DarkYellow
    Write-Host "  │          PCはベッドフレーム付近ですか？ │" -ForegroundColor DarkYellow
    Write-Host "  └─────────────────────────────────────────┘" -ForegroundColor DarkYellow
    $roomName = Read-Host "  居室名を入力（終了: q）"
    if ($roomName -eq "q" -or $roomName -eq "Q") { break }
    if ($roomName -eq "") {
        Write-Host "  ※ 居室名を入力してください" -ForegroundColor Yellow
        continue
    }

    $roomIndex++
    Write-Host ""
    Write-Host "  [$roomIndex] 「$roomName」を測定中（スキャンベース）..." -ForegroundColor Cyan

    $measurements  = @()
    $roomBand      = ""
    $roomChannel   = 0
    $roomBSSID     = ""
    $roomRadioType = ""
    $lastScan      = @()

    for ($i = 1; $i -le $MEASURE_COUNT; $i++) {
        $scanResult = Get-NearbyNetworks
        $lastScan   = $scanResult
        $targetAPs  = @($scanResult | Where-Object { $_.SSID -eq $targetSSID } | Sort-Object RSSI -Descending)

        if ($targetAPs.Count -gt 0) {
            # ① スキャン結果に対象SSIDが見つかった場合（ローミング影響なし）
            $best = $targetAPs[0]
            $rssi = $best.RSSI
            $measurements += $rssi
            if ($roomBand -eq "" -and $best.Band -ne "") {
                $roomBand      = $best.Band
                $roomChannel   = $best.Channel
                $roomBSSID     = $best.BSSID
                $roomRadioType = $best.RadioType
            }
            $rating = Get-SignalRating $rssi
            $bar    = "█" * $i + "░" * ($MEASURE_COUNT - $i)
            Write-Host "`r  [$bar] $i/$MEASURE_COUNT  RSSI: ${rssi} dBm（$rating）  " -NoNewline
        } else {
            # ② スキャン未検出 → 接続中APの信号値をフォールバックとして使用
            $conn = Get-WiFiInfo
            if ($conn.SSID -eq $targetSSID -and $conn.RSSI -gt -100) {
                $rssi = $conn.RSSI
                $measurements += $rssi
                if ($roomBand -eq "" -and $conn.Band -ne "") {
                    $roomBand      = $conn.Band
                    $roomChannel   = $conn.Channel
                    $roomBSSID     = $conn.BSSID
                    $roomRadioType = $conn.RadioType
                }
                $rating = Get-SignalRating $rssi
                $bar    = "█" * $i + "░" * ($MEASURE_COUNT - $i)
                Write-Host "`r  [$bar] $i/$MEASURE_COUNT  RSSI: ${rssi} dBm（$rating）[接続情報]  " -NoNewline
            } else {
                $measurements += -100
                $bar = "█" * $i + "░" * ($MEASURE_COUNT - $i)
                Write-Host "`r  [$bar] $i/$MEASURE_COUNT  圏外（$targetSSID 未検出）  " -NoNewline
            }
        }

        if ($i -lt $MEASURE_COUNT) { Start-Sleep -Seconds $MEASURE_INTERVAL }
    }

    # 統計計算（圏外値を除く）
    $validMeas = @($measurements | Where-Object { $_ -gt -100 })
    if ($validMeas.Count -gt 0) {
        $avg = [math]::Round(($validMeas | Measure-Object -Average).Average, 1)
        $min = ($validMeas | Measure-Object -Minimum).Minimum
        $max = ($validMeas | Measure-Object -Maximum).Maximum
    } else {
        $avg = -100; $min = -100; $max = -100
    }
    $avgRating = Get-SignalRating ([int]$avg)

    Write-Host ""
    Write-Host ""
    Write-Host "  ┌─────────────────────────────────┐" -ForegroundColor Green
    Write-Host "  │  結果: 平均 $avg dBm（$avgRating）" -ForegroundColor Green
    Write-Host "  │  最小: ${min} dBm  最大: ${max} dBm" -ForegroundColor Green
    Write-Host "  │  周辺AP検出数: $($lastScan.Count) 台（チャネル詳細はHTMLレポート参照）" -ForegroundColor Green
    Write-Host "  └─────────────────────────────────┘" -ForegroundColor Green
    Write-Host ""

    $allResults += [PSCustomObject]@{
        No             = $roomIndex
        RoomName       = $roomName
        SSID           = $targetSSID
        Band           = $roomBand
        Channel        = $roomChannel
        RadioType      = $roomRadioType
        AvgRSSI        = $avg
        MinRSSI        = $min
        MaxRSSI        = $max
        Rating         = $avgRating
        BSSID          = $roomBSSID
        Timestamp      = Get-Date -Format "HH:mm:ss"
        RawData        = ($measurements -join ",")
        NearbyNetworks = $lastScan
        Comment        = Get-RoomComment $avgRating
    }
}

# ----- 測定データなし -----
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
$dateStr   = Get-Date -Format "yyyy/MM/dd HH:mm"
$csvFile   = Join-Path $scriptDir "WiFi_Survey_${timestamp}.csv"
$htmlFile  = Join-Path $scriptDir "WiFi_Survey_${timestamp}.html"

# ----- CSV出力 -----
$csvHeader = "No,居室名,SSID,周波数帯,チャネル,無線規格,平均RSSI(dBm),最小(dBm),最大(dBm),評価,周辺AP検出数,BSSID,測定時刻,測定生データ(dBm)"
$csvLines  = @($csvHeader)
foreach ($r in $allResults) {
    $csvLines += "$($r.No),`"$($r.RoomName)`",`"$($r.SSID)`",$($r.Band),$($r.Channel),`"$($r.RadioType)`",$($r.AvgRSSI),$($r.MinRSSI),$($r.MaxRSSI),`"$($r.Rating)`",$($r.NearbyNetworks.Count),`"$($r.BSSID)`",$($r.Timestamp),`"$($r.RawData)`""
}
[System.IO.File]::WriteAllLines($csvFile, $csvLines, [System.Text.Encoding]::GetEncoding("shift_jis"))

# ----- 全体サマリー計算 -----
$overallAvg    = [math]::Round(($allResults | ForEach-Object { $_.AvgRSSI } | Measure-Object -Average).Average, 1)
$overallMin    = ($allResults | ForEach-Object { $_.MinRSSI } | Measure-Object -Minimum).Minimum
$overallRating = Get-SignalRating ([int]$overallAvg)
$weakRooms     = ($allResults | Where-Object { $_.AvgRSSI -lt -70 }).Count  # RSSI < -70dBm = 社内基準未達

# ----- 居室テーブルHTML -----
$tableRows = ""
foreach ($r in $allResults) {
    $color    = Get-SignalColor $r.Rating
    $barWidth = Get-BarWidth ([int]$r.AvgRSSI)
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
    <span style="font-weight:700;color:${color}">$($r.AvgRSSI) dBm</span>
  </div>
</td>
<td style="text-align:center;font-weight:600;color:${color}">$($r.MinRSSI) dBm</td>
<td style="text-align:center;font-weight:600;color:${color}">$($r.MaxRSSI) dBm</td>
<td style="text-align:center"><span style="color:${color};font-weight:700">$($r.Rating)</span></td>
<td style="text-align:center;font-size:12px;color:#94a3b8">$($r.Timestamp)</td>
</tr>
"@
}

$overallColor = Get-SignalColor $overallRating
$measurementNote = @"
<div style='background:#eff6ff;border:1.5px solid #93c5fd;border-radius:10px;padding:14px 20px;margin-bottom:20px'>
  <div style='font-size:13px;font-weight:700;color:#1e40af;margin-bottom:8px'>&#128204; 測定条件・評価基準</div>
  <div style='display:flex;gap:24px;flex-wrap:wrap;font-size:12px;color:#1e3a5f;line-height:1.8'>
    <div>
      <div style='font-weight:600;margin-bottom:2px'>【測定条件】</div>
      <div>・居室の扉を閉めた状態で測定</div>
      <div>・介護ベッドのフレーム付近（スリープセンサ設置想定位置）で測定</div>
    </div>
    <div>
      <div style='font-weight:600;margin-bottom:2px'>【社内目安値】</div>
      <div>・スリープセンサゲートウェイ設置場所での RSSI が <strong style='color:#1d4ed8'>-70 dBm 以上</strong> であることを目安としています</div>
      <div style='color:#64748b;font-size:11px;margin-top:2px'>※ RSSI はノイズ（干渉波）を考慮した値ではありません。RSSIが目安値を満たす場合でも、環境によっては接続が不安定になる場合があります。詳細な評価はNW担当者にご確認ください。</div>
    </div>
  </div>
</div>
"@

$weakWarning  = ""
if ($weakRooms -gt 0) {
    $weakWarning = "<div style='background:#fef2f2;border:2px solid #ef4444;border-radius:8px;padding:12px 16px;margin-top:16px;color:#991b1b;font-weight:600'>&#9888; 社内目安値（-70 dBm）を下回る居室が ${weakRooms} 箇所あります。NW担当者への確認事項としてお伝えください。</div>"
}

# ----- 居室別 周辺ネットワーク詳細HTML（アコーディオン形式）-----
$roomDetailHtml = ""
foreach ($r in $allResults) {

    # 周辺ネットワーク行を生成
    $nwRows = ""
    $nwIdx  = 0
    foreach ($nw in ($r.NearbyNetworks | Sort-Object Band, { -($_.RSSI) })) {
        $nwIdx++
        $nwColor = Get-SignalColor (Get-SignalRating $nw.RSSI)
        $nwBar   = Get-BarWidth $nw.RSSI
        $nwRows += @"
<tr>
<td style="text-align:center;font-size:11px">$nwIdx</td>
<td><strong>$(if($nw.SSID){"$($nw.SSID)"}else{"（非公開）"})</strong></td>
<td style="font-size:10px;font-family:monospace">$($nw.BSSID)</td>
<td style="text-align:center"><span style="background:$(if($nw.Band -eq '5GHz'){'#dbeafe;color:#2563eb'}else{'#fef3c7;color:#d97706'});padding:2px 6px;border-radius:4px;font-size:10px;font-weight:600">$($nw.Band)</span></td>
<td style="text-align:center">$($nw.Channel)</td>
<td>
  <div style="display:flex;align-items:center;gap:4px">
    <div style="width:60px;height:12px;background:#e5e7eb;border-radius:6px;overflow:hidden">
      <div style="width:${nwBar}%;height:100%;background:${nwColor};border-radius:6px"></div>
    </div>
    <span style="font-size:11px;font-weight:600;color:${nwColor}">$($nw.RSSI) dBm</span>
  </div>
</td>
<td style="font-size:10px">$($nw.Auth)</td>
</tr>
"@
    }

    # この居室でのチャネル分析
    $rCA     = Get-ChannelAnalysis $r.NearbyNetworks
    $ch24Htm = ""
    foreach ($ch in @(1,2,3,4,5,6,7,8,9,10,11,12,13)) {
        $info = $rCA.Channels24[$ch]
        $cnt  = $info.Count
        $bgC  = if ($cnt -eq 0) { "#d1fae5" } elseif ($cnt -le 2) { "#fef3c7" } elseif ($cnt -le 5) { "#fed7aa" } else { "#fecaca" }
        $fgC  = if ($cnt -eq 0) { "#059669" } elseif ($cnt -le 2) { "#d97706" } elseif ($cnt -le 5) { "#ea580c" } else { "#dc2626" }
        $hl   = if ($ch -eq 1 -or $ch -eq 6 -or $ch -eq 11) { "font-weight:700;border:2px solid $fgC;" } else { "" }
        $tgt  = if ($r.Band -eq "2.4GHz" -and $r.Channel -eq $ch) { "outline:3px solid #7c3aed;outline-offset:2px;" } else { "" }
        $ch24Htm += "<div style='text-align:center;padding:6px 3px;background:$bgC;border-radius:6px;min-width:40px;$hl$tgt' title='$($info.Networks -join ", ")'><div style='font-size:9px;color:#64748b'>ch$ch</div><div style='font-size:16px;font-weight:700;color:$fgC'>$cnt</div></div>"
    }
    $ch5Htm = ""
    foreach ($ch in @(36,40,44,48,52,56,60,64,100,104,108,112,116,120,124,128,132,136,140,149,153,157,161,165)) {
        $info = $rCA.Channels5[$ch]
        if (-not $info) { continue }
        $cnt  = $info.Count
        $bgC  = if ($cnt -eq 0) { "#d1fae5" } elseif ($cnt -le 1) { "#fef3c7" } else { "#fecaca" }
        $fgC  = if ($cnt -eq 0) { "#059669" } elseif ($cnt -le 1) { "#d97706" } else { "#dc2626" }
        $tgt  = if ($r.Band -eq "5GHz" -and $r.Channel -eq $ch) { "outline:3px solid #7c3aed;outline-offset:2px;" } else { "" }
        $ch5Htm += "<div style='text-align:center;padding:5px 2px;background:$bgC;border-radius:4px;min-width:32px;$tgt' title='$($info.Networks -join ", ")'><div style='font-size:8px;color:#64748b'>$ch</div><div style='font-size:12px;font-weight:700;color:$fgC'>$cnt</div></div>"
    }

    $c = $r.Comment   # Get-RoomComment の結果
    $roomDetailHtml += @"
<div class="room-detail">
<div class="room-detail-header" onclick="toggleDetail($($r.No))">
  <span>&#128246; [$($r.No)] $($r.RoomName) &nbsp;—&nbsp; 測定時刻: $($r.Timestamp) &nbsp;|&nbsp; 周辺AP検出数: $($r.NearbyNetworks.Count)</span>
  <span style="display:flex;align-items:center;gap:8px;flex-shrink:0">
    <span style="background:$($c.VerdictBg);color:$($c.VerdictColor);border:1px solid $($c.VerdictBorder);border-radius:999px;padding:2px 12px;font-size:12px;font-weight:700">$($c.Verdict)</span>
    <span class="toggle-icon" id="icon-$($r.No)">&#9660;</span>
  </span>
</div>
<div class="room-detail-body" id="body-$($r.No)" style="display:none">
  <div style="background:$($c.VerdictBg);border:1.5px solid $($c.VerdictBorder);border-radius:10px;padding:14px 18px;margin-bottom:16px">
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:6px">
      <span style="background:$($c.VerdictColor);color:#fff;border-radius:999px;padding:2px 14px;font-size:13px;font-weight:700">$($c.Verdict)</span>
      <span style="font-size:13px;font-weight:600;color:#1e293b">NW担当者への確認事項</span>
    </div>
    <p style="font-size:13px;color:#374151;line-height:1.7;margin:0">$($c.Signal)</p>
    <p style="font-size:11px;color:#94a3b8;margin-top:8px;margin-bottom:0">※ 上記は計測値に基づく事実の記録です。詳細な評価はNW担当者にご確認ください。</p>
  </div>
  <h4 style="font-size:13px;font-weight:700;margin-bottom:10px;color:#475569">周辺ネットワーク一覧（この地点・このタイミングでのスキャン結果）</h4>
  <table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:16px">
  <tr>
    <th style="padding:6px 8px;background:#f8fafc;color:#64748b;font-size:10px;font-weight:600;text-transform:uppercase;border-bottom:2px solid #e2e8f0">No</th>
    <th style="padding:6px 8px;background:#f8fafc;color:#64748b;font-size:10px;font-weight:600;text-transform:uppercase;border-bottom:2px solid #e2e8f0">SSID</th>
    <th style="padding:6px 8px;background:#f8fafc;color:#64748b;font-size:10px;font-weight:600;text-transform:uppercase;border-bottom:2px solid #e2e8f0">BSSID</th>
    <th style="padding:6px 8px;background:#f8fafc;color:#64748b;font-size:10px;font-weight:600;text-transform:uppercase;border-bottom:2px solid #e2e8f0">帯域</th>
    <th style="padding:6px 8px;background:#f8fafc;color:#64748b;font-size:10px;font-weight:600;text-transform:uppercase;border-bottom:2px solid #e2e8f0">Ch</th>
    <th style="padding:6px 8px;background:#f8fafc;color:#64748b;font-size:10px;font-weight:600;text-transform:uppercase;border-bottom:2px solid #e2e8f0">RSSI</th>
    <th style="padding:6px 8px;background:#f8fafc;color:#64748b;font-size:10px;font-weight:600;text-transform:uppercase;border-bottom:2px solid #e2e8f0">認証</th>
  </tr>
  $nwRows
  </table>
  <h4 style="font-size:13px;font-weight:700;margin-bottom:6px;color:#475569">チャネル利用状況（この地点でのスキャン結果）</h4>
  <p style="font-size:11px;color:#94a3b8;margin-bottom:10px">※ 各チャネルで検出されたAP台数を示します。NW担当者への参考情報としてお伝えください。当社ではチャネル干渉の良否判定は行いません。</p>
  <div style="margin-bottom:12px">
    <div style="font-size:11px;color:#64748b;margin-bottom:6px">2.4GHz帯 &nbsp;（太枠 = 非重複チャネル ch1/6/11 &nbsp;|&nbsp; 紫枠 = この居室で使用中のチャネル）</div>
    <div style="display:flex;gap:3px;flex-wrap:wrap">$ch24Htm</div>
  </div>
  <div>
    <div style="font-size:11px;color:#64748b;margin-bottom:6px">5GHz帯 &nbsp;（紫枠 = この居室で使用中のチャネル）</div>
    <div style="display:flex;gap:2px;flex-wrap:wrap">$ch5Htm</div>
  </div>
</div>
</div>
"@
}

# ----- HTMLレポート本体 -----
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
.container { max-width:1200px; margin:0 auto; }
.header { background:linear-gradient(135deg,#7c3aed,#5b21b6); color:#fff; padding:24px 32px; border-radius:12px; margin-bottom:20px; }
.header h1 { font-size:22px; font-weight:700; }
.header p  { font-size:13px; opacity:0.85; margin-top:4px; }
.summary { display:flex; gap:16px; margin-bottom:20px; flex-wrap:wrap; }
.summary-card { flex:1; min-width:180px; background:#fff; border-radius:12px; padding:20px; box-shadow:0 2px 8px rgba(0,0,0,0.08); text-align:center; }
.summary-card .label { font-size:12px; color:#64748b; font-weight:600; text-transform:uppercase; }
.summary-card .value { font-size:28px; font-weight:700; margin-top:4px; }
.summary-card .sub   { font-size:13px; color:#94a3b8; margin-top:2px; }
.card { background:#fff; border-radius:12px; box-shadow:0 2px 8px rgba(0,0,0,0.08); margin-bottom:20px; overflow:hidden; }
.card-header { padding:16px 24px; border-bottom:1px solid #e5e7eb; }
.card-header h2 { font-size:16px; font-weight:700; color:#1e293b; }
table { width:100%; border-collapse:collapse; }
th { text-align:left; padding:10px 12px; background:#f8fafc; color:#64748b; font-size:11px; font-weight:600; text-transform:uppercase; letter-spacing:0.5px; border-bottom:2px solid #e2e8f0; }
td { padding:10px 12px; font-size:13px; border-bottom:1px solid #f1f5f9; }
.legend { display:flex; gap:20px; padding:16px 24px; flex-wrap:wrap; border-top:1px solid #e5e7eb; }
.legend-item { display:flex; align-items:center; gap:6px; font-size:12px; color:#64748b; }
.legend-dot  { width:12px; height:12px; border-radius:50%; }
.note   { font-size:12px; color:#64748b; padding:12px 24px; background:#f8fafc; border-top:1px solid #e5e7eb; }
.footer { text-align:center; padding:16px; color:#94a3b8; font-size:12px; }
.btn-group { display:flex; gap:12px; justify-content:center; padding:24px; flex-wrap:wrap; }
.btn { padding:12px 32px; border-radius:8px; font-size:14px; font-weight:600; cursor:pointer; border:none; transition:all 0.2s; }
.btn-primary   { background:#7c3aed; color:#fff; }
.btn-primary:hover   { background:#6d28d9; }
.btn-secondary { background:#f1f5f9; color:#475569; }
.btn-secondary:hover { background:#e2e8f0; }
.room-detail { border:1px solid #e2e8f0; border-radius:8px; margin-bottom:10px; overflow:hidden; }
.room-detail-header { display:flex; justify-content:space-between; align-items:center; padding:12px 16px; background:#f8fafc; cursor:pointer; font-size:13px; font-weight:600; color:#1e293b; user-select:none; }
.room-detail-header:hover { background:#f1f5f9; }
.room-detail-body { padding:16px; }
.toggle-icon { font-size:12px; color:#94a3b8; flex-shrink:0; margin-left:8px; }
@media print {
  body{background:#fff;padding:10px;}
  .btn-group{display:none;}
  .summary-card,.card{box-shadow:none;border:1px solid #e5e7eb;}
  .room-detail-body{display:block !important;}
}
</style>
</head>
<body>
<div class="container">

<div class="header">
<h1>&#128225; Wi-Fi電波測定レポート</h1>
<p>施設名: $facilityName ／ 測定対象SSID: $targetSSID ／ 測定日時: $dateStr ／ 測定居室数: $($allResults.Count)</p>
</div>

<div class="summary">
<div class="summary-card">
<div class="label">全体平均</div>
<div class="value" style="color:$overallColor">${overallAvg} dBm</div>
<div class="sub">$overallRating</div>
</div>
<div class="summary-card">
<div class="label">最低値</div>
<div class="value" style="color:#ef4444">${overallMin} dBm</div>
<div class="sub">全居室の最小値</div>
</div>
<div class="summary-card">
<div class="label">測定居室数</div>
<div class="value" style="color:#2563eb">$($allResults.Count)</div>
<div class="sub">各${MEASURE_COUNT}回スキャン</div>
</div>
<div class="summary-card">
<div class="label">確認が必要な居室数</div>
<div class="value" style="color:$(if($weakRooms -gt 0){'#ef4444'}else{'#10b981'})">$weakRooms</div>
<div class="sub">社内目安値 -70 dBm 未達</div>
</div>
</div>

$measurementNote

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
<th>平均 RSSI</th>
<th style="text-align:center">最小</th>
<th style="text-align:center">最大</th>
<th style="text-align:center">評価</th>
<th style="text-align:center">時刻</th>
</tr>
$tableRows
</table>
<div class="legend">
<div class="legend-item"><div class="legend-dot" style="background:#10b981"></div> 優良（-65 dBm 以上）</div>
<div class="legend-item"><div class="legend-dot" style="background:#3b82f6"></div> 良好（-70 dBm 以上）</div>
<div class="legend-item"><div class="legend-dot" style="background:#f59e0b"></div> 普通（-80 dBm 以上）</div>
<div class="legend-item"><div class="legend-dot" style="background:#ef4444"></div> 弱い（-90 dBm 以上）</div>
<div class="legend-item"><div class="legend-dot" style="background:#991b1b"></div> 非常に弱い（-90 dBm 未満）</div>
</div>
<div class="note">
&#8251; 信号強度は netsh wlan show networks（スキャン結果）を RSSI(dBm) に変換して表示しています。移動中のローミング状態には依存しません。<br>
&#8251; 各居室 ${MEASURE_COUNT} 回スキャンの統計値です。チャネル利用状況はNW担当者への参考情報として居室詳細セクションに記載しています。
</div>
</div>

<div class="card">
<div class="card-header"><h2>&#128202; 居室別 周辺ネットワーク詳細</h2></div>
<div style="padding:16px 24px">
<p style="font-size:12px;color:#64748b;margin-bottom:16px">
各居室での測定時にスキャンした周辺ネットワークとチャネル利用状況です。各居室を展開して確認してください。
</p>
$roomDetailHtml
</div>
</div>

<div class="btn-group">
<button class="btn btn-primary" onclick="window.print()">&#128424; 印刷 / PDF保存</button>
<button class="btn btn-secondary" onclick="location.reload()">&#128260; 閉じる</button>
</div>

<div class="footer">
ライフリズムナビ Wi-Fi電波測定ツール v2.0 &mdash; エコナビスタ株式会社
</div>

</div>
<script>
function toggleDetail(n) {
  var body = document.getElementById('body-' + n);
  var icon = document.getElementById('icon-' + n);
  if (body.style.display === 'none') {
    body.style.display = 'block';
    icon.innerHTML = '&#9650;';
  } else {
    body.style.display = 'none';
    icon.innerHTML = '&#9660;';
  }
}
</script>
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
Write-Host "  CSV:  $csvFile" -ForegroundColor White
Write-Host "  HTML: $htmlFile" -ForegroundColor White
Write-Host ""

Start-Process $htmlFile

Write-Host "  ブラウザでレポートが開きます。" -ForegroundColor Gray
Write-Host ""
Read-Host "  Enterキーで終了"
