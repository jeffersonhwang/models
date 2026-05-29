param(
  [Parameter(Mandatory = $false)]
  [string]$Path = "NOW.xlsx",

  [Parameter(Mandatory = $false)]
  [string]$CIK = "0001373715"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Get-QuarterEndDate {
  param([int]$Year, [int]$Quarter)
  switch ($Quarter) {
    1 { return Get-Date -Year $Year -Month 3 -Day 31 }
    2 { return Get-Date -Year $Year -Month 6 -Day 30 }
    3 { return Get-Date -Year $Year -Month 9 -Day 30 }
    4 { return Get-Date -Year $Year -Month 12 -Day 31 }
    default { throw "Invalid quarter: $Quarter" }
  }
}

function Get-QuarterStartDate {
  param([int]$Year, [int]$Quarter)
  switch ($Quarter) {
    1 { return Get-Date -Year $Year -Month 1 -Day 1 }
    2 { return Get-Date -Year $Year -Month 4 -Day 1 }
    3 { return Get-Date -Year $Year -Month 7 -Day 1 }
    4 { return Get-Date -Year $Year -Month 10 -Day 1 }
    default { throw "Invalid quarter: $Quarter" }
  }
}

function Quarter-Code {
  param([int]$Year, [int]$Quarter)
  return ("Q{0}{1:00}" -f $Quarter, ($Year % 100))
}

function Parse-QuarterCode {
  param([string]$Code)
  if ($Code -notmatch '^Q([1-4])(\d\d)$') { throw "Invalid quarter code: $Code" }
  $q = [int]$Matches[1]
  $yy = [int]$Matches[2]
  $year = if ($yy -ge 70) { 1900 + $yy } else { 2000 + $yy }
  return [pscustomobject]@{ Year = $year; Quarter = $q }
}

function Round-Millions {
  param([Nullable[double]]$Value)
  if ($null -eq $Value) { return $null }
  return [math]::Round($Value / 1e6, 0, [MidpointRounding]::AwayFromZero)
}

function Coalesce-Number {
  param(
    [Parameter(Mandatory = $false)][Nullable[double]]$Value,
    [Parameter(Mandatory = $true)][double]$Default
  )
  if ($null -eq $Value) { return $Default }
  return [double]$Value
}

function Excel-ColumnLetters {
  param([int]$Column)
  if ($Column -lt 1) { throw "Invalid column: $Column" }
  $letters = ""
  $n = $Column
  while ($n -gt 0) {
    $n--
    $letters = ([char]([int](65 + ($n % 26)))) + $letters
    $n = [math]::Floor($n / 26)
  }
  return $letters
}

function Excel-CellAddress {
  param([int]$Row, [int]$Column)
  return ("{0}{1}" -f (Excel-ColumnLetters -Column $Column), $Row)
}

$headers = @{
  "User-Agent" = "martinshkreli-models (contact: research@example.com)"
}

$instantUsdCache = @{}
$ytdUsdCache = @{}
$sharesCache = @{}

$cikPadded = $CIK.PadLeft(10, "0")
$uri = "https://data.sec.gov/api/xbrl/companyfacts/CIK{0}.json" -f $cikPadded
$cachePath = "sec_companyfacts_CIK{0}.json" -f $cikPadded
if (Test-Path $cachePath) {
  Write-Host "Loading cached SEC companyfacts: $cachePath"
  $companyFacts = Get-Content -Raw -Path $cachePath | ConvertFrom-Json
} else {
  Write-Host "Downloading SEC companyfacts: $uri"
  $companyFacts = $null
  for ($i = 1; $i -le 3; $i++) {
    try {
      $companyFacts = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
      break
    }
    catch {
      if ($i -ge 3) { throw }
      Write-Host "Download failed (attempt $i). Retrying..."
      Start-Sleep -Seconds (2 * $i)
    }
  }
}
$gaap = $companyFacts.facts.'us-gaap'

function Get-Units {
  param(
    [Parameter(Mandatory = $true)][string]$Tag,
    [Parameter(Mandatory = $true)][string]$Unit
  )
  $p = $gaap.$Tag
  if ($null -eq $p) { return @() }
  $u = $p.units.$Unit
  if ($null -eq $u) { return @() }
  return @($u)
}

function Pick-InstantUSD {
  param(
    [Parameter(Mandatory = $true)][string]$Tag,
    [Parameter(Mandatory = $true)][datetime]$End
  )
  $endS = $End.ToString("yyyy-MM-dd")
  $key = "$Tag|$endS"
  if ($instantUsdCache.ContainsKey($key)) { return $instantUsdCache[$key] }

  $facts = Get-Units -Tag $Tag -Unit "USD"
  if (-not $facts -or $facts.Count -eq 0) {
    $instantUsdCache[$key] = $null
    return $null
  }

  $best = $null
  $bestFiled = [datetime]"1900-01-01"
  foreach ($f in $facts) {
    if ($f.end -ne $endS) { continue }
    if ($null -ne $f.start -and $f.start -ne "") { continue }
    $filed = if ($null -ne $f.filed -and $f.filed -ne "") { [datetime]$f.filed } else { [datetime]"1900-01-01" }
    if ($filed -gt $bestFiled) { $best = $f; $bestFiled = $filed }
  }

  $val = if ($null -eq $best) { $null } else { [double]$best.val }
  $instantUsdCache[$key] = $val
  return $val
}

function Pick-YTDUSD {
  param(
    [Parameter(Mandatory = $true)][string]$Tag,
    [Parameter(Mandatory = $true)][datetime]$End,
    [Parameter(Mandatory = $true)][string]$FormRegex
  )
  $startS = (Get-Date -Year $End.Year -Month 1 -Day 1).ToString("yyyy-MM-dd")
  $endS = $End.ToString("yyyy-MM-dd")
  $formKey = if ($FormRegex -match "10-Q") { "10-Q" } elseif ($FormRegex -match "10-K") { "10-K" } else { $FormRegex }
  $key = "$Tag|$formKey|$startS|$endS"
  if ($ytdUsdCache.ContainsKey($key)) { return $ytdUsdCache[$key] }

  $facts = Get-Units -Tag $Tag -Unit "USD"
  if (-not $facts -or $facts.Count -eq 0) {
    $ytdUsdCache[$key] = $null
    return $null
  }

  $best = $null
  $bestFiled = [datetime]"1900-01-01"
  foreach ($f in $facts) {
    if ($f.start -ne $startS) { continue }
    if ($f.end -ne $endS) { continue }
    if ($null -eq $f.form -or ($f.form -notmatch $FormRegex)) { continue }
    $filed = if ($null -ne $f.filed -and $f.filed -ne "") { [datetime]$f.filed } else { [datetime]"1900-01-01" }
    if ($filed -gt $bestFiled) { $best = $f; $bestFiled = $filed }
  }

  $val = if ($null -eq $best) { $null } else { [double]$best.val }
  $ytdUsdCache[$key] = $val
  return $val
}

function Get-QuarterUSDFromYTD {
  param(
    [Parameter(Mandatory = $true)][string]$Tag,
    [Parameter(Mandatory = $true)][int]$Year,
    [Parameter(Mandatory = $true)][int]$Quarter
  )
  $end = Get-QuarterEndDate -Year $Year -Quarter $Quarter
  if ($Quarter -lt 4) {
    $ytd = Pick-YTDUSD -Tag $Tag -End $end -FormRegex '^10-Q'
    if ($null -eq $ytd) { return $null }
    if ($Quarter -eq 1) { return $ytd }
    $prevEnd = Get-QuarterEndDate -Year $Year -Quarter ($Quarter - 1)
    $prevYtd = Pick-YTDUSD -Tag $Tag -End $prevEnd -FormRegex '^10-Q'
    if ($null -eq $prevYtd) { return $null }
    return $ytd - $prevYtd
  }

  $fy = Pick-YTDUSD -Tag $Tag -End $end -FormRegex '^10-K'
  if ($null -eq $fy) { return $null }
  $q3End = Get-QuarterEndDate -Year $Year -Quarter 3
  $q3Ytd = Pick-YTDUSD -Tag $Tag -End $q3End -FormRegex '^10-Q'
  if ($null -eq $q3Ytd) { return $null }
  return $fy - $q3Ytd
}

function Pick-QuarterShares {
  param(
    [Parameter(Mandatory = $true)][string]$Tag,
    [Parameter(Mandatory = $true)][datetime]$Start,
    [Parameter(Mandatory = $true)][datetime]$End,
    [Parameter(Mandatory = $true)][string]$FormRegex
  )
  $startS = $Start.ToString("yyyy-MM-dd")
  $endS = $End.ToString("yyyy-MM-dd")
  $formKey = if ($FormRegex -match "10-Q") { "10-Q" } elseif ($FormRegex -match "10-K") { "10-K" } else { $FormRegex }
  $key = "$Tag|$formKey|$startS|$endS"
  if ($sharesCache.ContainsKey($key)) { return $sharesCache[$key] }

  $facts = Get-Units -Tag $Tag -Unit "shares"
  if (-not $facts -or $facts.Count -eq 0) {
    $sharesCache[$key] = $null
    return $null
  }

  $best = $null
  $bestFiled = [datetime]"1900-01-01"
  foreach ($f in $facts) {
    if ($f.start -ne $startS) { continue }
    if ($f.end -ne $endS) { continue }
    if ($null -eq $f.form -or ($f.form -notmatch $FormRegex)) { continue }
    $filed = if ($null -ne $f.filed -and $f.filed -ne "") { [datetime]$f.filed } else { [datetime]"1900-01-01" }
    if ($filed -gt $bestFiled) { $best = $f; $bestFiled = $filed }
  }

  $val = if ($null -eq $best) { $null } else { [double]$best.val }
  $sharesCache[$key] = $val
  return $val
}

function Get-QuarterDilutedShares {
  param([int]$Year, [int]$Quarter)
  $tag = "WeightedAverageNumberOfDilutedSharesOutstanding"
  $end = Get-QuarterEndDate -Year $Year -Quarter $Quarter
  if ($Quarter -lt 4) {
    $start = Get-QuarterStartDate -Year $Year -Quarter $Quarter
    return Pick-QuarterShares -Tag $tag -Start $start -End $end -FormRegex '^10-Q'
  }

  $fyStart = Get-Date -Year $Year -Month 1 -Day 1
  $fy = Pick-QuarterShares -Tag $tag -Start $fyStart -End $end -FormRegex '^10-K'
  if ($null -eq $fy) { return $null }
  $q1 = Get-QuarterDilutedShares -Year $Year -Quarter 1
  $q2 = Get-QuarterDilutedShares -Year $Year -Quarter 2
  $q3 = Get-QuarterDilutedShares -Year $Year -Quarter 3
  if ($null -eq $q1 -or $null -eq $q2 -or $null -eq $q3) { return $null }
  return 4 * $fy - ($q1 + $q2 + $q3)
}

function Get-PretaxUSD {
  param([int]$Year, [int]$Quarter)
  $tag1 = "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest"
  $tag2 = "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"
  $v = Get-QuarterUSDFromYTD -Tag $tag1 -Year $Year -Quarter $Quarter
  if ($null -ne $v) { return $v }
  return Get-QuarterUSDFromYTD -Tag $tag2 -Year $Year -Quarter $Quarter
}

function Get-FXUSD {
  param([int]$Year, [int]$Quarter)
  $tag1 = "EffectOfExchangeRateOnCashAndCashEquivalents"
  $tag2 = "EffectOfExchangeRateOnCashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"
  $v = Get-QuarterUSDFromYTD -Tag $tag1 -Year $Year -Quarter $Quarter
  if ($null -ne $v) { return $v }
  return Get-QuarterUSDFromYTD -Tag $tag2 -Year $Year -Quarter $Quarter
}

function Get-CashTotalComponents {
  param([datetime]$End)
  $cash = Pick-InstantUSD -Tag "CashAndCashEquivalentsAtCarryingValue" -End $End
  if ($null -eq $cash) { return $null }
  $st = Pick-InstantUSD -Tag "AvailableForSaleSecuritiesDebtSecuritiesCurrent" -End $End
  $lt = Pick-InstantUSD -Tag "AvailableForSaleSecuritiesDebtSecuritiesNoncurrent" -End $End
  return [pscustomobject]@{
    Cash = $cash
    AfsCurrent = (Coalesce-Number -Value $st -Default 0)
    AfsNoncurrent = (Coalesce-Number -Value $lt -Default 0)
    Total = ($cash + (Coalesce-Number -Value $st -Default 0) + (Coalesce-Number -Value $lt -Default 0))
  }
}

function Get-CommissionsComponents {
  param([datetime]$End)
  $cur = Pick-InstantUSD -Tag "ContractWithCustomerAssetNetCurrent" -End $End
  $non = Pick-InstantUSD -Tag "ContractWithCustomerAssetNetNoncurrent" -End $End
  if ($null -eq $cur -and $null -eq $non) { return $null }
  return [pscustomobject]@{
    Current = (Coalesce-Number -Value $cur -Default 0)
    Noncurrent = (Coalesce-Number -Value $non -Default 0)
    Total = ((Coalesce-Number -Value $cur -Default 0) + (Coalesce-Number -Value $non -Default 0))
  }
}

function Get-DeferredRevenueComponents {
  param([datetime]$End)
  $cur = Pick-InstantUSD -Tag "ContractWithCustomerLiabilityCurrent" -End $End
  $non = Pick-InstantUSD -Tag "ContractWithCustomerLiabilityNoncurrent" -End $End
  if ($null -ne $cur -or $null -ne $non) {
    return [pscustomobject]@{
      Current = (Coalesce-Number -Value $cur -Default 0)
      Noncurrent = (Coalesce-Number -Value $non -Default 0)
      Total = ((Coalesce-Number -Value $cur -Default 0) + (Coalesce-Number -Value $non -Default 0))
    }
  }
  $cur2 = Pick-InstantUSD -Tag "DeferredRevenueCurrent" -End $End
  $non2 = Pick-InstantUSD -Tag "DeferredRevenueNoncurrent" -End $End
  if ($null -ne $cur2 -or $null -ne $non2) {
    return [pscustomobject]@{
      Current = (Coalesce-Number -Value $cur2 -Default 0)
      Noncurrent = (Coalesce-Number -Value $non2 -Default 0)
      Total = ((Coalesce-Number -Value $cur2 -Default 0) + (Coalesce-Number -Value $non2 -Default 0))
    }
  }
  $tot = Pick-InstantUSD -Tag "DeferredRevenue" -End $End
  if ($null -eq $tot) { return $null }
  return [pscustomobject]@{ Current = $tot; Noncurrent = 0; Total = $tot }
}

function Get-LeaseLiabilityComponents {
  param([datetime]$End)
  $tot = Pick-InstantUSD -Tag "OperatingLeaseLiability" -End $End
  $cur = Pick-InstantUSD -Tag "OperatingLeaseLiabilityCurrent" -End $End
  $non = Pick-InstantUSD -Tag "OperatingLeaseLiabilityNoncurrent" -End $End
  if ($null -ne $cur -or $null -ne $non) {
    return [pscustomobject]@{
      Current = (Coalesce-Number -Value $cur -Default 0)
      Noncurrent = (Coalesce-Number -Value $non -Default 0)
      Total = ((Coalesce-Number -Value $cur -Default 0) + (Coalesce-Number -Value $non -Default 0))
    }
  }
  if ($null -ne $tot) { return [pscustomobject]@{ Current = $tot; Noncurrent = 0; Total = $tot } }
  return $null
}

function Get-DebtTotal {
  param([datetime]$End)
  $conv = Pick-InstantUSD -Tag "ConvertibleLongTermNotesPayable" -End $End
  $convCur = Pick-InstantUSD -Tag "ConvertibleNotesPayableCurrent" -End $End
  if ($null -ne $conv -or $null -ne $convCur) {
    return ((Coalesce-Number -Value $conv -Default 0) + (Coalesce-Number -Value $convCur -Default 0))
  }
  $lt = Pick-InstantUSD -Tag "LongTermDebtNoncurrent" -End $End
  $ltc = Pick-InstantUSD -Tag "LongTermDebtCurrent" -End $End
  if ($null -ne $lt -or $null -ne $ltc) {
    return ((Coalesce-Number -Value $lt -Default 0) + (Coalesce-Number -Value $ltc -Default 0))
  }
  $cd = Pick-InstantUSD -Tag "ConvertibleDebt" -End $End
  if ($null -ne $cd) { return $cd }
  return $null
}

function Get-Prepaid {
  param([datetime]$End)
  foreach ($t in @("OtherPrepaidExpenseCurrent", "PrepaidExpenseAndOtherAssetsCurrent")) {
    $v = Pick-InstantUSD -Tag $t -End $End
    if ($null -ne $v) { return $v }
  }
  return $null
}

function Get-DeferredTaxAssets {
  param([datetime]$End)
  foreach ($t in @("DeferredIncomeTaxAssetsNet", "DeferredTaxAssetsNet", "DeferredTaxAssetsNetCurrent", "DeferredIncomeTaxesAndOtherAssetsNoncurrent")) {
    $v = Pick-InstantUSD -Tag $t -End $End
    if ($null -ne $v) { return $v }
  }
  return $null
}

function Get-ServicesRevenueUSD {
  param([int]$Year, [int]$Quarter)
  foreach ($t in @("SalesRevenueServicesNet", "TechnologyServicesRevenue")) {
    $v = Get-QuarterUSDFromYTD -Tag $t -Year $Year -Quarter $Quarter
    if ($null -ne $v) { return $v }
  }
  return 0
}

function Latest-ReportedQuarterEnd {
  $facts = Get-Units -Tag "RevenueFromContractWithCustomerExcludingAssessedTax" -Unit "USD"
  $qFacts = $facts | Where-Object { $_.form -eq "10-Q" -and $_.end -and $_.start -and $_.fp -match '^Q[1-3]$' }
  if (-not $qFacts -or $qFacts.Count -eq 0) { throw "Unable to find latest 10-Q revenue facts." }
  return ($qFacts | Sort-Object { [datetime]$_.end } -Descending | Select-Object -First 1).end
}

$latest10QEndS = Latest-ReportedQuarterEnd
$latest10QEnd = [datetime]$latest10QEndS
$latestYear = $latest10QEnd.Year
$latestQuarter = switch ($latest10QEnd.Month) { 3 { 1 } 6 { 2 } 9 { 3 } default { throw "Unexpected latest 10-Q end date: $latest10QEndS" } }

Write-Host "Latest reported 10-Q quarter end: $latest10QEndS (Q$latestQuarter $latestYear)"

$startYear = 2020
$quartersToFill = @()
for ($y = $startYear; $y -le $latestYear; $y++) {
  for ($q = 1; $q -le 4; $q++) {
    if ($y -eq $latestYear -and $q -gt $latestQuarter) { continue }
    $quartersToFill += (Quarter-Code -Year $y -Quarter $q)
  }
}

$q425 = (Quarter-Code -Year 2025 -Quarter 4)
if ($quartersToFill -notcontains $q425) { $quartersToFill += $q425 }

Write-Host "Opening workbook: $Path"
$absPath = (Resolve-Path $Path).Path
$excel = $null
$wb = $null
$ws = $null
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.ScreenUpdating = $false
  $excel.EnableEvents = $false
  try { $excel.AutomationSecurity = 3 } catch {} # msoAutomationSecurityForceDisable

  $wb = $excel.Workbooks.Open($absPath, 0, $false)
  for ($i = 0; $i -lt 160; $i++) {
    try {
      if ($excel.Ready) { break }
    } catch {}
    Start-Sleep -Milliseconds 250
  }
  $ws = $wb.Worksheets.Item("Model")
  for ($i = 0; $i -lt 120; $i++) {
    try {
      $null = $ws.Range("C2").Value2
      break
    } catch {
      Start-Sleep -Milliseconds 500
    }
  }

  $headerRow = 2
  $quarterToCol = @{}
  $maxScan = 200
  $maxCol = 0
  for ($c = 1; $c -le $maxScan; $c++) {
    $addr = Excel-CellAddress -Row $headerRow -Column $c
    $v = $ws.Range($addr).Value2
    if ($null -ne $v -and ($v -is [string]) -and $v -match '^Q[1-4]\d\d$') {
      $quarterToCol[$v] = $c
      if ($c -gt $maxCol) { $maxCol = $c }
    }
  }
  if ($maxCol -eq 0) { throw "No quarter headers found in row $headerRow." }

  foreach ($code in $quartersToFill) {
    if ($quarterToCol.ContainsKey($code)) { continue }
    $newCol = $maxCol + 1
    $ws.Range((Excel-CellAddress -Row $headerRow -Column $newCol)).Value2 = $code
    $quarterToCol[$code] = $newCol
    $maxCol = $newCol
  }

  $rows = @{
    Subscription = 11
    Services = 12
    COGS = 14
    SM = 16
    RD = 17
    GA = 18
    Interest = 21
    Taxes = 23
    Shares = 26
    Cash = 30
    AR = 31
    Commissions = 32
    Prepaid = 33
    PPE = 34
    LeaseAsset = 35
    Intangibles = 36
    TaxAssets = 37
    OtherAssets = 38
    AP = 41
    Accrued = 42
    DR = 43
    LeaseLiab = 44
    Debt = 45
    OtherLiab = 46
    SE = 47
    ReportedNI = 52
    DA = 53
    AmortDC = 54
    SBC = 55
    CFO_Taxes = 56
    CFO_Other = 57
    WC = 58
    CapEx = 61
    Acq = 62
    Invest = 63
    Converts = 66
    ESOP = 67
    CFF_Taxes = 68
    FX = 70
  }

  foreach ($code in $quartersToFill) {
    if ($code -eq "Q425") { continue } # leave Q4'25 blank per request
    $qInfo = Parse-QuarterCode -Code $code
    $y = $qInfo.Year
    $q = $qInfo.Quarter
    $end = Get-QuarterEndDate -Year $y -Quarter $q
    $col = $quarterToCol[$code]

    $rev = Get-QuarterUSDFromYTD -Tag "RevenueFromContractWithCustomerExcludingAssessedTax" -Year $y -Quarter $q
    $cogs = Get-QuarterUSDFromYTD -Tag "CostOfRevenue" -Year $y -Quarter $q
    $opexSM = Get-QuarterUSDFromYTD -Tag "SellingAndMarketingExpense" -Year $y -Quarter $q
    $opexRD = Get-QuarterUSDFromYTD -Tag "ResearchAndDevelopmentExpense" -Year $y -Quarter $q
    $opexGA = Get-QuarterUSDFromYTD -Tag "GeneralAndAdministrativeExpense" -Year $y -Quarter $q
    $opInc = Get-QuarterUSDFromYTD -Tag "OperatingIncomeLoss" -Year $y -Quarter $q
    $pretax = Get-PretaxUSD -Year $y -Quarter $q
    $tax = Get-QuarterUSDFromYTD -Tag "IncomeTaxExpenseBenefit" -Year $y -Quarter $q
    $interest = if ($null -eq $pretax -or $null -eq $opInc) { $null } else { $pretax - $opInc }

    $services = Get-ServicesRevenueUSD -Year $y -Quarter $q
    if ($null -eq $rev) { throw "Missing revenue for $code" }
    $subscription = $rev - (Coalesce-Number -Value $services -Default 0)

    $shares = Get-QuarterDilutedShares -Year $y -Quarter $q

    $cashComponents = Get-CashTotalComponents -End $end
    $ar = Pick-InstantUSD -Tag "AccountsReceivableNetCurrent" -End $end
    $commComponents = Get-CommissionsComponents -End $end
    $prepaid = Get-Prepaid -End $end
    $ppe = Pick-InstantUSD -Tag "PropertyPlantAndEquipmentNet" -End $end
    $leaseAsset = Pick-InstantUSD -Tag "OperatingLeaseRightOfUseAsset" -End $end
    $intangibles = Pick-InstantUSD -Tag "IntangibleAssetsNetExcludingGoodwill" -End $end
    $taxAssets = Get-DeferredTaxAssets -End $end
    $assets = Pick-InstantUSD -Tag "Assets" -End $end

    $ap = Pick-InstantUSD -Tag "AccountsPayableCurrent" -End $end
    $accrued = Pick-InstantUSD -Tag "AccruedLiabilitiesCurrent" -End $end
    $drComponents = Get-DeferredRevenueComponents -End $end
    $leaseLiabComponents = Get-LeaseLiabilityComponents -End $end
    $debt = Get-DebtTotal -End $end
    $liab = Pick-InstantUSD -Tag "Liabilities" -End $end
    $se = Pick-InstantUSD -Tag "StockholdersEquity" -End $end

    $ni = Get-QuarterUSDFromYTD -Tag "NetIncomeLoss" -Year $y -Quarter $q
    $da = Get-QuarterUSDFromYTD -Tag "DepreciationDepletionAndAmortization" -Year $y -Quarter $q
    $amortDC = Get-QuarterUSDFromYTD -Tag "AmortizationOfDeferredSalesCommissions" -Year $y -Quarter $q
    $sbc = Get-QuarterUSDFromYTD -Tag "ShareBasedCompensation" -Year $y -Quarter $q
    $cfo = Get-QuarterUSDFromYTD -Tag "NetCashProvidedByUsedInOperatingActivities" -Year $y -Quarter $q
    $capex = Get-QuarterUSDFromYTD -Tag "PaymentsToAcquirePropertyPlantAndEquipment" -Year $y -Quarter $q
    $acq = Get-QuarterUSDFromYTD -Tag "PaymentsToAcquireBusinessesNetOfCashAcquired" -Year $y -Quarter $q
    $cfi = Get-QuarterUSDFromYTD -Tag "NetCashProvidedByUsedInInvestingActivities" -Year $y -Quarter $q
    $cff = Get-QuarterUSDFromYTD -Tag "NetCashProvidedByUsedInFinancingActivities" -Year $y -Quarter $q
    $fx = Get-FXUSD -Year $y -Quarter $q

    $amortDC = (Coalesce-Number -Value $amortDC -Default 0)
    $acq = (Coalesce-Number -Value $acq -Default 0)

    $wcResidual = if ($null -eq $cfo -or $null -eq $ni -or $null -eq $da -or $null -eq $sbc) { $null } else { $cfo - ($ni + $da + $amortDC + $sbc) }
    $invResidual = if ($null -eq $cfi -or $null -eq $capex) { $null } else { $cfi - ($capex + $acq) }

    $cashM = Round-Millions $cashComponents.Cash
    $stM = Round-Millions $cashComponents.AfsCurrent
    $ltM = Round-Millions $cashComponents.AfsNoncurrent

    $commTotal = if ($null -ne $commComponents) { $commComponents.Total } else { $null }
    $assetsKnown = @(
      $cashComponents.Total,
      $ar,
      $commTotal,
      $prepaid,
      $ppe,
      $leaseAsset,
      $intangibles,
      $taxAssets
    )
    $assetsKnownSum = ($assetsKnown | Where-Object { $null -ne $_ } | Measure-Object -Sum).Sum
    $otherAssets = if ($null -eq $assets) { $null } else { $assets - $assetsKnownSum }

    $drTotal = if ($null -ne $drComponents) { $drComponents.Total } else { $null }
    $leaseLiabTotal = if ($null -ne $leaseLiabComponents) { $leaseLiabComponents.Total } else { $null }
    $liabKnown = @(
      $ap,
      $accrued,
      $drTotal,
      $leaseLiabTotal,
      $debt
    )
    $liabKnownSum = ($liabKnown | Where-Object { $null -ne $_ } | Measure-Object -Sum).Sum
    $otherLiab = if ($null -eq $liab) { $null } else { $liab - $liabKnownSum }

    $ws.Range((Excel-CellAddress -Row $rows.Subscription -Column $col)).Value2 = (Round-Millions $subscription)
    $ws.Range((Excel-CellAddress -Row $rows.Services -Column $col)).Value2 = (Round-Millions $services)
    $ws.Range((Excel-CellAddress -Row $rows.COGS -Column $col)).Value2 = (Round-Millions $cogs)
    $ws.Range((Excel-CellAddress -Row $rows.SM -Column $col)).Value2 = (Round-Millions $opexSM)
    $ws.Range((Excel-CellAddress -Row $rows.RD -Column $col)).Value2 = (Round-Millions $opexRD)
    $ws.Range((Excel-CellAddress -Row $rows.GA -Column $col)).Value2 = (Round-Millions $opexGA)
    $ws.Range((Excel-CellAddress -Row $rows.Interest -Column $col)).Value2 = (Round-Millions $interest)
    $ws.Range((Excel-CellAddress -Row $rows.Taxes -Column $col)).Value2 = (Round-Millions $tax)
    $ws.Range((Excel-CellAddress -Row $rows.Shares -Column $col)).Value2 = if ($null -eq $shares) { $null } else { [math]::Round($shares / 1e6, 0, [MidpointRounding]::AwayFromZero) }

    $ws.Range((Excel-CellAddress -Row $rows.Cash -Column $col)).Formula = "=$cashM+$stM+$ltM"
    $ws.Range((Excel-CellAddress -Row $rows.AR -Column $col)).Value2 = (Round-Millions $ar)

    if ($null -ne $commComponents) {
      $commCurM = Round-Millions $commComponents.Current
      $commNonM = Round-Millions $commComponents.Noncurrent
      $ws.Range((Excel-CellAddress -Row $rows.Commissions -Column $col)).Formula = "=$commCurM+$commNonM"
    } else {
      $ws.Range((Excel-CellAddress -Row $rows.Commissions -Column $col)).Value2 = $null
    }

    $ws.Range((Excel-CellAddress -Row $rows.Prepaid -Column $col)).Value2 = (Round-Millions $prepaid)
    $ws.Range((Excel-CellAddress -Row $rows.PPE -Column $col)).Value2 = (Round-Millions $ppe)
    $ws.Range((Excel-CellAddress -Row $rows.LeaseAsset -Column $col)).Value2 = (Round-Millions $leaseAsset)
    $ws.Range((Excel-CellAddress -Row $rows.Intangibles -Column $col)).Value2 = (Round-Millions $intangibles)
    $ws.Range((Excel-CellAddress -Row $rows.TaxAssets -Column $col)).Value2 = (Round-Millions $taxAssets)
    $ws.Range((Excel-CellAddress -Row $rows.OtherAssets -Column $col)).Value2 = (Round-Millions $otherAssets)

    $ws.Range((Excel-CellAddress -Row $rows.AP -Column $col)).Value2 = (Round-Millions $ap)
    $ws.Range((Excel-CellAddress -Row $rows.Accrued -Column $col)).Value2 = (Round-Millions $accrued)

    if ($null -ne $drComponents) {
      $drCurM = Round-Millions $drComponents.Current
      $drNonM = Round-Millions $drComponents.Noncurrent
      $ws.Range((Excel-CellAddress -Row $rows.DR -Column $col)).Formula = "=$drCurM+$drNonM"
    } else {
      $ws.Range((Excel-CellAddress -Row $rows.DR -Column $col)).Value2 = $null
    }

    if ($null -ne $leaseLiabComponents) {
      $llCurM = Round-Millions $leaseLiabComponents.Current
      $llNonM = Round-Millions $leaseLiabComponents.Noncurrent
      $ws.Range((Excel-CellAddress -Row $rows.LeaseLiab -Column $col)).Formula = "=$llCurM+$llNonM"
    } else {
      $ws.Range((Excel-CellAddress -Row $rows.LeaseLiab -Column $col)).Value2 = $null
    }

    $ws.Range((Excel-CellAddress -Row $rows.Debt -Column $col)).Value2 = (Round-Millions $debt)
    $ws.Range((Excel-CellAddress -Row $rows.OtherLiab -Column $col)).Value2 = (Round-Millions $otherLiab)
    $ws.Range((Excel-CellAddress -Row $rows.SE -Column $col)).Value2 = (Round-Millions $se)

    $ws.Range((Excel-CellAddress -Row $rows.ReportedNI -Column $col)).Value2 = (Round-Millions $ni)
    $ws.Range((Excel-CellAddress -Row $rows.DA -Column $col)).Value2 = (Round-Millions $da)
    $ws.Range((Excel-CellAddress -Row $rows.AmortDC -Column $col)).Value2 = (Round-Millions $amortDC)
    $ws.Range((Excel-CellAddress -Row $rows.SBC -Column $col)).Value2 = (Round-Millions $sbc)
    $ws.Range((Excel-CellAddress -Row $rows.CFO_Taxes -Column $col)).Value2 = 0
    $ws.Range((Excel-CellAddress -Row $rows.CFO_Other -Column $col)).Value2 = 0
    $ws.Range((Excel-CellAddress -Row $rows.WC -Column $col)).Value2 = (Round-Millions $wcResidual)

    $ws.Range((Excel-CellAddress -Row $rows.CapEx -Column $col)).Value2 = (Round-Millions $capex)
    $ws.Range((Excel-CellAddress -Row $rows.Acq -Column $col)).Value2 = (Round-Millions $acq)
    $ws.Range((Excel-CellAddress -Row $rows.Invest -Column $col)).Value2 = (Round-Millions $invResidual)

    $ws.Range((Excel-CellAddress -Row $rows.Converts -Column $col)).Value2 = 0
    $ws.Range((Excel-CellAddress -Row $rows.ESOP -Column $col)).Value2 = (Round-Millions $cff)
    $ws.Range((Excel-CellAddress -Row $rows.CFF_Taxes -Column $col)).Value2 = 0
    $ws.Range((Excel-CellAddress -Row $rows.FX -Column $col)).Value2 = (Round-Millions $fx)
  }

  $wb.Save()
  Write-Host "Saved: $Path"
}
finally {
  if ($null -ne $ws) {
    try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) } catch {}
  }
  if ($null -ne $wb) {
    try { $wb.Close($false) | Out-Null } catch {}
    try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) } catch {}
  }
  if ($null -ne $excel) {
    try { $excel.Quit() | Out-Null } catch {}
    try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) } catch {}
  }
}

Write-Host ""
Write-Host "Validations (millions):"

function FY-USD {
  param([string]$Tag, [int]$Year)
  $end = Get-QuarterEndDate -Year $Year -Quarter 4
  $fy = Pick-YTDUSD -Tag $Tag -End $end -FormRegex '^10-K'
  return $fy
}

foreach ($y in $startYear..$latestYear) {
  if ($y -gt 2025) { continue }
  $qVals = @()
  foreach ($q in 1..4) {
    if ($y -eq $latestYear -and $q -gt $latestQuarter) { continue }
    $qVals += (Get-QuarterUSDFromYTD -Tag "RevenueFromContractWithCustomerExcludingAssessedTax" -Year $y -Quarter $q)
  }
  if ($qVals.Count -lt 4 -and $y -ge 2025) { continue }
  $sumQ = ($qVals | Measure-Object -Sum).Sum
  $fy = FY-USD -Tag "RevenueFromContractWithCustomerExcludingAssessedTax" -Year $y
  if ($null -ne $fy) {
    $diffM = (Round-Millions($fy - $sumQ))
    Write-Host ("Revenue {0}: FY {1} vs sum(Q) {2} diff {3}" -f $y, (Round-Millions $fy), (Round-Millions $sumQ), $diffM)
  }

  $opQ = @()
  foreach ($q in 1..4) {
    if ($y -eq $latestYear -and $q -gt $latestQuarter) { continue }
    $opQ += (Get-QuarterUSDFromYTD -Tag "OperatingIncomeLoss" -Year $y -Quarter $q)
  }
  $opSum = ($opQ | Measure-Object -Sum).Sum
  $opFY = FY-USD -Tag "OperatingIncomeLoss" -Year $y
  if ($null -ne $opFY) {
    $diffM = (Round-Millions($opFY - $opSum))
    Write-Host ("OpInc  {0}: FY {1} vs sum(Q) {2} diff {3}" -f $y, (Round-Millions $opFY), (Round-Millions $opSum), $diffM)
  }
}

Write-Host ""
Write-Host ("Spot-check latest quarter revenue/opinc: {0} revenue={1} opInc={2}" -f (Quarter-Code -Year $latestYear -Quarter $latestQuarter), (Round-Millions (Get-QuarterUSDFromYTD -Tag 'RevenueFromContractWithCustomerExcludingAssessedTax' -Year $latestYear -Quarter $latestQuarter)), (Round-Millions (Get-QuarterUSDFromYTD -Tag 'OperatingIncomeLoss' -Year $latestYear -Quarter $latestQuarter)))
