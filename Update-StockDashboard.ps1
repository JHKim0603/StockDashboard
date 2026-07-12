<#
  Update-StockDashboard.ps1
  Fetches live quotes (Yahoo Finance) + recent headlines (Google News RSS) for the tickers below
  and regenerates dashboard.html.
  Run manually by double-clicking run.bat, or right-click > Run with PowerShell.
#>

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot

# Add/remove tickers here.
#   DisplayName : optional, leave $null to use the exchange's long name
#   NewsQuery   : what to search Google News for (use the underlying company name for derivative products)
#   NewsLang    : "ko" or "en" — controls which Google News edition is queried
#   IsIndex     : $true for market indices — suppresses the currency symbol and shows "지수" as the market chip
$tickers = @(
    [PSCustomObject]@{ Symbol = "005930.KS"; DisplayName = "삼성전자";    MarketLabel = "KOSPI";  NewsQuery = "삼성전자";               NewsLang = "ko"; IsIndex = $false },
    [PSCustomObject]@{ Symbol = "000660.KS"; DisplayName = "SK하이닉스";  MarketLabel = "KOSPI";  NewsQuery = "SK하이닉스";             NewsLang = "ko"; IsIndex = $false },
    [PSCustomObject]@{ Symbol = "SNDK";      DisplayName = $null;        MarketLabel = "NASDAQ"; NewsQuery = "SanDisk SNDK stock";    NewsLang = "en"; IsIndex = $false },
    [PSCustomObject]@{ Symbol = "MU";        DisplayName = $null;        MarketLabel = "NASDAQ"; NewsQuery = "Micron Technology MU stock"; NewsLang = "en"; IsIndex = $false },
    [PSCustomObject]@{ Symbol = "^IXIC";     DisplayName = "나스닥";      MarketLabel = "지수";    NewsQuery = "Nasdaq Composite index"; NewsLang = "en"; IsIndex = $true },
    [PSCustomObject]@{ Symbol = "^GSPC";     DisplayName = "S&P 500";    MarketLabel = "지수";    NewsQuery = "S&P 500 index";          NewsLang = "en"; IsIndex = $true },
    [PSCustomObject]@{ Symbol = "^KS11";     DisplayName = "KOSPI";      MarketLabel = "지수";    NewsQuery = "코스피 지수";             NewsLang = "ko"; IsIndex = $true }
)

$currencySymbols = @{ KRW = "₩"; USD = "$"; }
$newsLocales = @{
    ko = @{ hl = "ko";    gl = "KR"; ceid = "KR:ko" }
    en = @{ hl = "en-US"; gl = "US"; ceid = "US:en" }
}
$headers = @{ "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36" }

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
            [PSCustomObject]@{
                title  = $title
                source = $source
                date   = $pubDate.ToString("yyyy-MM-dd")
                link   = $it.link
            }
        })
    } catch {
        Write-Warning "News fetch failed for '$query': $($_.Exception.Message)"
        @()
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

    [PSCustomObject]@{
        name      = $name
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
    }
}

Write-Host "Fetching live quotes and headlines..."
$stocks = foreach ($t in $tickers) {
    Write-Host "  - $($t.Symbol)"
    Get-StockSnapshot $t
    Start-Sleep -Milliseconds 400  # be gentle with Yahoo's unofficial endpoint across 7 tickers
}

$stocksJson = ConvertTo-Json -InputObject @($stocks) -Depth 6
$fetchedAt = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")

$template = Get-Content -Path (Join-Path $root "template.html") -Raw -Encoding UTF8
$output = $template.Replace("__STOCKS_JSON__", $stocksJson).Replace("__FETCHED_AT__", $fetchedAt)

$outPath = Join-Path $root "dashboard.html"
Set-Content -Path $outPath -Value $output -Encoding UTF8

Write-Host "Dashboard updated: $outPath"
Start-Process $outPath
