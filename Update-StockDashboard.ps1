<#
  Update-StockDashboard.ps1
  Fetches live quotes (Yahoo Finance) + recent headlines (Google News RSS) for the tickers below
  and regenerates dashboard.html.
  Run manually by double-clicking run.bat, or right-click > Run with PowerShell.
#>

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot

# Tickers live in watchlist.json, not here — only "Symbol" is required, e.g. { "Symbol": "AAPL" }.
# Everything else (DisplayName, MarketLabel, NewsQuery, NewsLang, FinanceMode/Code) is optional and
# auto-derived from the symbol below; set any of them explicitly in watchlist.json to override.
#   Symbol suffix convention: ".KS" = KOSPI, ".KQ" = KOSDAQ, "^" prefix = index, anything else = NASDAQ
#   (NYSE tickers need an explicit "FinanceCode": "SYMBOL.N" override — the auto-default assumes NASDAQ)
function Resolve-TickerConfig {
    param($raw)

    $symbol = $raw.Symbol
    if (-not $symbol) { throw "watchlist.json 항목에 Symbol이 없습니다: $($raw | ConvertTo-Json -Compress)" }

    $isIndex = if ($null -ne $raw.IsIndex) { [bool]$raw.IsIndex } else { $symbol.StartsWith("^") }
    $isDomestic = $symbol.EndsWith(".KS") -or $symbol.EndsWith(".KQ")

    $marketLabel =
        if ($raw.MarketLabel) { $raw.MarketLabel }
        elseif ($isIndex) { "지수" }
        elseif ($symbol.EndsWith(".KS")) { "KOSPI" }
        elseif ($symbol.EndsWith(".KQ")) { "KOSDAQ" }
        else { "NASDAQ" }

    $newsLang = if ($raw.NewsLang) { $raw.NewsLang } elseif ($isDomestic) { "ko" } else { "en" }

    $financeMode =
        if ($raw.FinanceMode) { $raw.FinanceMode }
        elseif ($isIndex) { $null }
        elseif ($isDomestic) { "domestic" }
        else { "overseas" }

    $financeCode =
        if ($raw.FinanceCode) { $raw.FinanceCode }
        elseif ($financeMode -eq "domestic") { $symbol -replace '\.KS$|\.KQ$', '' }
        elseif ($financeMode -eq "overseas") { "$symbol.O" }
        else { $null }

    [PSCustomObject]@{
        Symbol      = $symbol
        DisplayName = $raw.DisplayName
        MarketLabel = $marketLabel
        NewsQuery   = $raw.NewsQuery
        NewsLang    = $newsLang
        IsIndex     = $isIndex
        FinanceMode = $financeMode
        FinanceCode = $financeCode
    }
}

$watchlistRaw = Get-Content -Path (Join-Path $root "watchlist.json") -Raw -Encoding UTF8 | ConvertFrom-Json
$tickers = @($watchlistRaw | ForEach-Object { Resolve-TickerConfig $_ })

$currencySymbols = @{ KRW = "₩"; USD = "$"; }
$newsLocales = @{
    ko = @{ hl = "ko";    gl = "KR"; ceid = "KR:ko" }
    en = @{ hl = "en-US"; gl = "US"; ceid = "US:en" }
}
$headers = @{ "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36" }

function Get-KoreanTranslation {
    param($text)

    try {
        $uri = "https://translate.googleapis.com/translate_a/single?client=gtx&sl=en&tl=ko&dt=t&q=" + [uri]::EscapeDataString($text)
        $resp = Invoke-WebRequest -Uri $uri -Headers $headers -UseBasicParsing
        $parsed = $resp.Content | ConvertFrom-Json
        # Google splits long input into several segments under $parsed[0]; each segment's
        # translated text is element [0] — join them back into one string.
        (($parsed[0] | ForEach-Object { $_[0] }) -join "").Trim()
    } catch {
        Write-Warning "Translation failed for '$text': $($_.Exception.Message)"
        $null
    }
}

function Get-NewsHeadlines {
    param($query, $lang, $max = 4)

    $loc = $newsLocales[$lang]
    $uri = "https://news.google.com/rss/search?q=" + [uri]::EscapeDataString($query) + "&hl=$($loc.hl)&gl=$($loc.gl)&ceid=$($loc.ceid)"
    try {
        $raw = Invoke-WebRequest -Uri $uri -Headers $headers -UseBasicParsing
        [xml]$rss = $raw.Content
        $items = $rss.rss.channel.item | Select-Object -First $max
        @(foreach ($it in $items) {
            $title = $it.title
            $source = $null
            # Google News sometimes appends the outlet name twice (native + romanized), e.g. "... - 조선비즈 - Chosunbiz"
            for ($k = 0; $k -lt 2; $k++) {
                if ($title -match '^(.*) - ([^-]{1,40})$') {
                    $title = $matches[1].Trim(); $source = $matches[2].Trim()
                } else { break }
            }
            $pubDate = [System.DateTimeOffset]::Parse($it.pubDate)

            $titleKo = $null
            if ($lang -eq "en") {
                $translated = Get-KoreanTranslation $title
                if ($translated -and $translated -ne $title) { $titleKo = $translated }
                Start-Sleep -Milliseconds 200  # be gentle with the unofficial translate endpoint
            }

            [PSCustomObject]@{
                title         = if ($titleKo) { $titleKo } else { $title }
                titleOriginal = if ($titleKo) { $title } else { $null }
                source        = $source
                date          = $pubDate.ToString("yyyy-MM-dd")
                link          = $it.link
            }
        })
    } catch {
        Write-Warning "News fetch failed for '$query': $($_.Exception.Message)"
        @()
    }
}

function Get-FinanceSnapshot {
    param($mode, $code)

    if (-not $mode -or -not $code) { return $null }

    try {
        # Quarterly, not annual: annual periods are ~1 year apart and stale for most of the
        # year; quarterly gives 5-6 recent reporting periods, which is what's actually useful here.
        if ($mode -eq "domestic") {
            $uri = "https://m.stock.naver.com/api/stock/$code/finance/quarter"
            $resp = Invoke-RestMethod -Uri $uri -Headers $headers
            $fi = $resp.financeInfo
            if (-not $fi -or -not $fi.trTitleList -or $fi.trTitleList.Count -eq 0) { return $null }
            $periods = $fi.trTitleList
            $rows = $fi.rowList
            $summaryParts = @($resp.corporationSummary.comment1, $resp.corporationSummary.comment2) | Where-Object { $_ }
            $summary = if ($summaryParts) { $summaryParts -join " " } else { $null }
            $unit = "억원"
        } else {
            $uri = "https://api.stock.naver.com/stock/$code/finance/quarter"
            $resp = Invoke-RestMethod -Uri $uri -Headers $headers
            if (-not $resp.trTitleList -or $resp.trTitleList.Count -eq 0) { return $null }
            $periods = $resp.trTitleList
            $rows = $resp.rowList
            $summary = $null
            $unit = $resp.unit
        }

        $wantedTitles = @("매출액", "영업이익", "EBIT", "당기순이익", "EPS")
        $orderedPeriods = @($periods | Sort-Object key)
        $today = Get-Date

        $rowsOut = foreach ($rowTitle in $wantedTitles) {
            $row = $rows | Where-Object { $_.title -eq $rowTitle } | Select-Object -First 1
            if (-not $row) { continue }
            [PSCustomObject]@{
                title  = if ($rowTitle -eq "EBIT") { "영업이익(EBIT)" } else { $rowTitle }
                values = @(foreach ($p in $orderedPeriods) {
                    $col = $row.columns.($p.key)
                    if ($col -and $col.value) { $col.value } else { "-" }
                })
            }
        }
        if (-not $rowsOut) { return $null }

        $periodsOut = foreach ($p in $orderedPeriods) {
            $isEstimate = $false
            if ($mode -eq "domestic") {
                # Domestic quarters flag consensus (not-yet-reported) periods explicitly and reliably.
                $isEstimate = ($p.isConsensus -eq "Y")
            } else {
                # Overseas quarters don't set isConsensus reliably; a period whose end date hasn't
                # happened yet is necessarily a forecast, so fall back to a date comparison.
                [DateTime]$parsedEnd = Get-Date
                if ([DateTime]::TryParse($p.key, [ref]$parsedEnd)) { $isEstimate = $parsedEnd -gt $today }
            }
            [PSCustomObject]@{ label = $p.title; isEstimate = $isEstimate }
        }

        $nextEstimateLabel = ($periodsOut | Where-Object { $_.isEstimate } | Select-Object -Last 1).label

        # Valuation band (PER/PBR trend + min/avg/max over the same quarters already fetched
        # above — no extra request). Naver's overseas quarterly data only has PBR, not PER.
        $valuation = foreach ($metric in @("PER", "PBR")) {
            $row = $rows | Where-Object { $_.title -eq $metric } | Select-Object -First 1
            if (-not $row) { continue }
            $points = for ($i = 0; $i -lt $orderedPeriods.Count; $i++) {
                $p = $orderedPeriods[$i]
                $col = $row.columns.($p.key)
                $val = if ($col -and $col.value -and $col.value -ne "-") { [double]($col.value -replace ',', '') } else { $null }
                [PSCustomObject]@{ label = $p.title; isEstimate = $periodsOut[$i].isEstimate; value = $val }
            }
            $actualValues = @($points | Where-Object { -not $_.isEstimate -and $null -ne $_.value } | ForEach-Object { $_.value })
            if ($actualValues.Count -lt 2) { continue }  # not enough history for a meaningful band
            [PSCustomObject]@{
                metric = $metric
                points = @($points)
                min    = ($actualValues | Measure-Object -Minimum).Minimum
                max    = ($actualValues | Measure-Object -Maximum).Maximum
                avg    = [math]::Round((($actualValues | Measure-Object -Average).Average), 2)
            }
        }

        [PSCustomObject]@{
            unit              = $unit
            periods           = @($periodsOut)
            rows              = @($rowsOut)
            summary           = $summary
            nextEstimateLabel = $nextEstimateLabel
            valuation         = @($valuation)
        }
    } catch {
        Write-Warning "Finance fetch failed for '$code' ($mode): $($_.Exception.Message)"
        $null
    }
}

function Get-ConsensusSnapshot {
    param($mode, $code)

    if (-not $mode -or -not $code) { return $null }

    try {
        $uri = if ($mode -eq "domestic") { "https://m.stock.naver.com/api/stock/$code/integration" } else { "https://api.stock.naver.com/stock/$code/integration" }
        $resp = Invoke-RestMethod -Uri $uri -Headers $headers
        $ci = $resp.consensusInfo
        if (-not $ci -or -not $ci.priceTargetMean) { return $null }

        # No free source publishes each brokerage's individual target price (that's paywalled
        # FnGuide/WiseFn data) — the best available breakdown is the underlying report list
        # (broker name + title + date) behind the consensus average. Domestic-only: Korean
        # brokerages don't publish research on foreign-listed names like MU/SNDK.
        $reports = $null
        if ($mode -eq "domestic") {
            try {
                $researchUri = "https://m.stock.naver.com/api/research/stock/$code"
                $researchResp = Invoke-RestMethod -Uri $researchUri -Headers $headers
                $reports = @($researchResp | Select-Object -First 6 | ForEach-Object {
                    [PSCustomObject]@{
                        broker = $_.brokerName
                        title  = $_.title
                        date   = $_.writeDate
                        link   = "https://finance.naver.com/research/company_read.naver?nid=$($_.researchId)"
                    }
                })
            } catch {
                Write-Warning "Research list fetch failed for '$code': $($_.Exception.Message)"
            }
        }

        [PSCustomObject]@{
            targetPrice = [double]($ci.priceTargetMean -replace ',', '')
            targetHigh  = if ($ci.priceTargetHigh) { [double]($ci.priceTargetHigh -replace ',', '') } else { $null }
            targetLow   = if ($ci.priceTargetLow)  { [double]($ci.priceTargetLow  -replace ',', '') } else { $null }
            recommScore = if ($ci.recommMean) { [double]$ci.recommMean } else { $null }
            asOf        = $ci.createDate
            reports     = $reports
        }
    } catch {
        Write-Warning "Consensus fetch failed for '$code' ($mode): $($_.Exception.Message)"
        $null
    }
}

function Get-StockSnapshot {
    param($cfg)

    $uri = "https://query1.finance.yahoo.com/v8/finance/chart/$($cfg.Symbol)?interval=1d&range=1y"
    $resp = Invoke-RestMethod -Uri $uri -Headers $headers
    $result = $resp.chart.result[0]
    $meta = $result.meta
    $closes = $result.indicators.quote[0].close
    $timestamps = $result.timestamp

    $pairs = for ($i = 0; $i -lt $closes.Count; $i++) {
        if ($null -ne $closes[$i]) {
            [PSCustomObject]@{
                Date  = [DateTimeOffset]::FromUnixTimeSeconds($timestamps[$i]).ToLocalTime().ToString("yyyy-MM-dd")
                Close = [math]::Round([double]$closes[$i], 2)
            }
        }
    }

    $fullSeries = @($pairs.Close)
    $fullDates  = @($pairs.Date)

    # Yahoo's meta.fiftyTwoWeekHigh/Low occasionally comes back as 0 for some tickers;
    # fall back to the min/max of the fetched 1y series so we never divide by zero.
    $rangeHigh = if ($meta.fiftyTwoWeekHigh -and $meta.fiftyTwoWeekHigh -gt 0) { $meta.fiftyTwoWeekHigh } else { ($fullSeries | Measure-Object -Maximum).Maximum }
    $rangeLow  = if ($meta.fiftyTwoWeekLow  -and $meta.fiftyTwoWeekLow  -gt 0) { $meta.fiftyTwoWeekLow }  else { ($fullSeries | Measure-Object -Minimum).Minimum }

    # trailing-20-session stats for the auto-summary sentence (independent of the chart's own zoom range)
    $recent = $pairs | Select-Object -Last 20
    $rSeries = @($recent.Close)
    $first = $rSeries[0]
    $lastVal = $rSeries[-1]
    $periodPct = (($lastVal - $first) / $first) * 100

    $upDays = 0; $downDays = 0
    for ($i = 1; $i -lt $rSeries.Count; $i++) {
        if ($rSeries[$i] -gt $rSeries[$i - 1]) { $upDays++ }
        elseif ($rSeries[$i] -lt $rSeries[$i - 1]) { $downDays++ }
    }

    $pctFromHigh = (($lastVal - $rangeHigh) / $rangeHigh) * 100
    $pctFromLow  = (($lastVal - $rangeLow)  / $rangeLow)  * 100

    $periodSign = if ($periodPct -ge 0) { "+" } else { "" }
    $lowSign    = if ($pctFromLow -ge 0) { "+" } else { "" }

    $summary = "최근 {0}거래일 중 {1}일 상승 · {2}일 하락, 기간 시작 대비 {3}{4:N1}%. 52주 고점 대비 {5:N1}%, 저점 대비 {6}{7:N1}% 수준입니다." -f `
        $rSeries.Count, $upDays, $downDays, $periodSign, $periodPct, $pctFromHigh, $lowSign, $pctFromLow

    $name = if ($cfg.DisplayName) { $cfg.DisplayName } else { $meta.longName }
    if ($cfg.IsIndex) {
        $currency = ""
        $market = $cfg.MarketLabel
    } else {
        $currency = $currencySymbols[$meta.currency]
        if (-not $currency) { $currency = "$($meta.currency) " }
        $market = $meta.currency
    }

    $newsQuery = if ($cfg.NewsQuery) { $cfg.NewsQuery } else { $name }
    $news = Get-NewsHeadlines -query $newsQuery -lang $cfg.NewsLang
    $finance = Get-FinanceSnapshot -mode $cfg.FinanceMode -code $cfg.FinanceCode
    $consensus = Get-ConsensusSnapshot -mode $cfg.FinanceMode -code $cfg.FinanceCode

    [PSCustomObject]@{
        name      = $name
        symbol    = $cfg.Symbol
        ticker    = "$($cfg.Symbol) · $($cfg.MarketLabel)"
        market    = $market
        currency  = $currency
        series    = $fullSeries
        dates     = $fullDates
        volume    = $meta.regularMarketVolume
        rangeLow  = $rangeLow
        rangeHigh = $rangeHigh
        summary   = $summary
        news      = $news
        finance   = $finance
        consensus = $consensus
    }
}

# Korea has no DST, so UTC+9 is always correct — avoids TimeZoneInfo id mismatches between
# Windows PowerShell 5.1 locally ("Korea Standard Time") and pwsh on the Linux Actions runner
# ("Asia/Seoul"), which previously left $fetchedAt on raw runner-local time (UTC in CI).
$nowKst = (Get-Date).ToUniversalTime().AddHours(9)

function Get-MovingAverage {
    param($series, $window)
    $out = New-Object 'object[]' $series.Count
    $sum = 0.0
    for ($i = 0; $i -lt $series.Count; $i++) {
        $sum += $series[$i]
        if ($i -ge $window) { $sum -= $series[$i - $window] }
        $out[$i] = if ($i -ge $window - 1) { $sum / $window } else { $null }
    }
    $out
}

function Get-CrossSignal {
    # Mirrors the dashboard's own detectCross() in template.html — same MA windows, same
    # 5-day lookback for an actual sign-flip event, not just "which MA is on top now".
    param($series, $lookback = 5)
    $ma20 = Get-MovingAverage -series $series -window 20
    $ma60 = Get-MovingAverage -series $series -window 60
    $diffs = [System.Collections.ArrayList]@()
    for ($i = $series.Count - 1; $i -ge 0 -and $diffs.Count -lt ($lookback + 1); $i--) {
        if ($null -eq $ma20[$i] -or $null -eq $ma60[$i]) { break }
        [void]$diffs.Insert(0, ($ma20[$i] - $ma60[$i]))
    }
    for ($i = 1; $i -lt $diffs.Count; $i++) {
        if ($diffs[$i - 1] -le 0 -and $diffs[$i] -gt 0) { return "golden" }
        if ($diffs[$i - 1] -ge 0 -and $diffs[$i] -lt 0) { return "dead" }
    }
    return $null
}

Write-Host "Fetching USD/KRW exchange rate..."
$usdKrw = $null
try {
    $fxResp = Invoke-RestMethod -Uri "https://query1.finance.yahoo.com/v8/finance/chart/KRW=X?interval=1d&range=5d" -Headers $headers
    $usdKrw = [math]::Round([double]$fxResp.chart.result[0].meta.regularMarketPrice, 2)
} catch {
    Write-Warning "Exchange rate fetch failed: $($_.Exception.Message)"
}

Write-Host "Fetching live quotes and headlines..."
$stocks = foreach ($t in $tickers) {
    Write-Host "  - $($t.Symbol)"
    Get-StockSnapshot $t
    Start-Sleep -Milliseconds 400  # be gentle with Yahoo's unofficial endpoint across 7 tickers
}

$stocksJson = ConvertTo-Json -InputObject @($stocks) -Depth 8
$usdKrwJson = ConvertTo-Json -InputObject $usdKrw
$fetchedAt = $nowKst.ToString("yyyy-MM-ddTHH:mm:ss") + "+09:00"  # $nowKst's DateTimeKind is still Utc after the manual +9h shift, so "zzz" would report +00:00 — append the known-fixed KST offset literally instead.

$template = Get-Content -Path (Join-Path $root "template.html") -Raw -Encoding UTF8
$output = $template.Replace("__STOCKS_JSON__", $stocksJson).Replace("__USDKRW_JSON__", $usdKrwJson).Replace("__FETCHED_AT__", $fetchedAt)

$outPath = Join-Path $root "dashboard.html"
Set-Content -Path $outPath -Value $output -Encoding UTF8

Write-Host "Dashboard updated: $outPath"
if (-not $env:CI) {
    Start-Process $outPath
}

# --- Email summary -----------------------------------------------------------------
# Generates email-summary.html + email-subject.txt every run (harmless, useful for local
# preview) — actually SENDING the email is a separate step, done only by the GitHub Actions
# workflow, so testing locally never spams the inbox.
Write-Host "Building email summary..."

$dayNames = @("일", "월", "화", "수", "목", "금", "토")
$emailDateStr = "{0}년 {1}월 {2}일 ({3})" -f $nowKst.Year, $nowKst.Month, $nowKst.Day, $dayNames[[int]$nowKst.DayOfWeek]

$rowsHtml = foreach ($s in $stocks) {
    $last = $s.series[-1]
    $prev = $s.series[-2]
    $diff = $last - $prev
    $pct = if ($prev -ne 0) { ($diff / $prev) * 100 } else { 0 }
    $up = $diff -ge 0
    $color = if ($up) { "#0ca30c" } else { "#e34948" }
    $arrow = if ($up) { "▲" } else { "▼" }
    $priceFmt = if ($s.currency -eq "₩") { "{0:N0}" -f $last } else { "{0:N2}" -f $last }

    $cross = Get-CrossSignal -series $s.series
    $crossTag =
        if ($cross -eq "golden") { " <span style='color:#0ca30c;font-weight:700;'>[골든크로스]</span>" }
        elseif ($cross -eq "dead") { " <span style='color:#e34948;font-weight:700;'>[데드크로스]</span>" }
        else { "" }

    $newsHtml = ""
    if ($s.news -and $s.news.Count -gt 0) {
        $newsLines = foreach ($n in ($s.news | Select-Object -First 2)) {
            "<div style='font-size:12px;color:#52514e;margin-top:3px;'>· <a href='$($n.link)' style='color:#2a78d6;text-decoration:none;'>$($n.title)</a></div>"
        }
        $newsHtml = $newsLines -join ""
    }

    @"
<tr>
  <td style="padding:10px 12px;border-bottom:1px solid #e1e0d9;">
    <div style="font-weight:600;font-size:13px;color:#0b0b0b;">$($s.name)$crossTag</div>
    <div style="font-size:11px;color:#898781;">$($s.ticker)</div>
    $newsHtml
  </td>
  <td style="padding:10px 12px;border-bottom:1px solid #e1e0d9;text-align:right;white-space:nowrap;vertical-align:top;">
    <div style="font-weight:650;font-size:14px;color:#0b0b0b;">$($s.currency)$priceFmt</div>
    <div style="font-size:12px;font-weight:600;color:$color;">$arrow $([math]::Abs($pct).ToString("N2"))%</div>
  </td>
</tr>
"@
}

$fxLine = if ($usdKrw) { " · USD/KRW $($usdKrw.ToString('N2'))" } else { "" }

$emailHtml = @"
<!DOCTYPE html>
<html><body style="margin:0;padding:0;background:#f9f9f7;font-family:'Malgun Gothic',sans-serif;">
<div style="max-width:600px;margin:0 auto;padding:24px 16px;">
  <h2 style="margin:0 0 4px;color:#0b0b0b;">JH 투자 DASHBOARD</h2>
  <div style="font-size:12px;color:#898781;margin-bottom:16px;">$emailDateStr 기준$fxLine</div>
  <table style="width:100%;border-collapse:collapse;background:#ffffff;border:1px solid #e1e0d9;border-radius:8px;">
    $($rowsHtml -join "`n")
  </table>
  <div style="margin-top:16px;font-size:11px;color:#898781;line-height:1.6;">
    시세 출처: Yahoo Finance. 뉴스 출처: Google 뉴스(영문 기사는 자동 번역). 투자 판단 참고용으로만 사용하세요 — 개인 용도 요약입니다.<br>
    실적·목표주가·전체 차트 등 자세한 내용은 대시보드에서 확인하세요.
  </div>
</div>
</body></html>
"@

# Set-Content -Encoding UTF8 writes a BOM, which would leak into the mail subject/body as a
# stray character in some clients — write both files BOM-less instead.
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText((Join-Path $root "email-summary.html"), $emailHtml, $utf8NoBom)
[System.IO.File]::WriteAllText((Join-Path $root "email-subject.txt"), "JH 투자 DASHBOARD 요약 - $emailDateStr", $utf8NoBom)
Write-Host "Email summary written: email-summary.html"
