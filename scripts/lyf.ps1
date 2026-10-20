# lyf.ps1 — Býr til gagnaskrá fyrir lyfjakostnaðarmælaborðið (key = lyf).
#
# Notkun:
#   powershell -ExecutionPolicy Bypass -File lyf.ps1 -Mappa "C:\...\Gögn\Október" -Manudur 2026-09
#
# Les (read-only, gegnum Excel COM):
#   8.1 Lyfjakaup ... Svið (KG samtals)       -> D.blend
#   8.2 Lyfjakaup ... Sviðseiningar (KG02–05) -> D.kg
# Blöð: "Gjaldabókun", "Yfirlit ATC flokka", "30 kostnaðarmestu samheiti lyfj..."
# Skrifar: <rót>\data\<Manudur>\lyf.js  (window.D = {...};)
#
# Fjöldi mánaða ræðst af línunni "Mánuðir til og með NN" í topp-30 blaðinu
# (til vara: síðasti mánuður með hreyfingu í samtölulínu).

param(
  [Parameter(Mandatory = $true)][string]$Mappa,
  [Parameter(Mandatory = $true)][string]$Manudur
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2

$Rot = Split-Path -Parent $PSScriptRoot
$Inv = [Globalization.CultureInfo]::InvariantCulture

# ── Fastar sem mælaborðið byggir á ─────────────────────────────────────────
$KG_DEF = [ordered]@{
  'KG02' = @{ name = 'KG02 Sérgreinar bráðaþj.'; color = '#7c3aed' }
  'KG03' = @{ name = 'KG03 Ferlideildir';        color = '#2563eb' }
  'KG04' = @{ name = 'KG04 Legudeildir';         color = '#0891b2' }
  'KG05' = @{ name = 'KG05 Endurhæfing';         color = '#059669' }
}
$EXP_DEF = [ordered]@{
  '535100' = 'Lyf'; '535110' = 'Hjúkrunarvörur'; '535120' = 'Næringarefni';
  '535130' = 'Snyrtivörur og vítamín'; '535140' = 'Ýmsar vörur'; '535800' = 'Innöndunarefni';
  '597120' = 'Birgðastýring'; '597130' = 'Lyfjaskömmtun'; '597140' = 'Umsýslugjald'; '597210' = 'Lyfjablöndun'
}
$MAN_STUTT = @('Jan','Feb','Mar','Apr','Maí','Jún','Júl','Ágú','Sep','Okt','Nóv','Des')
$MAN_LANGT = @('Janúar','Febrúar','Mars','Apríl','Maí','Júní','Júlí','Ágúst','September','Október','Nóvember','Desember')
$MAN_BIL   = @('Jan','Feb','Mars','Apríl','Maí','Júní','Júlí','Ágúst','Sept','Okt','Nóv','Des')

# ── Hjálparföll ────────────────────────────────────────────────────────────
function Get-Str($arr, [int]$r, [int]$c) {
  if ($c -gt $arr.GetLength(1)) { return '' }
  $v = $arr[$r, $c]
  if ($null -eq $v) { return '' }
  return ([string]$v)
}
function Get-Num($arr, [int]$r, [int]$c) {
  if ($c -lt 1 -or $c -gt $arr.GetLength(1)) { return $null }
  $v = $arr[$r, $c]
  if ($null -eq $v) { return $null }
  if ($v -is [double] -or $v -is [int] -or $v -is [decimal]) { return [double]$v }
  $s = ([string]$v).Trim()
  if ($s -eq '') { return $null }
  $d = 0.0
  if ([double]::TryParse($s, [Globalization.NumberStyles]::Float, $Inv, [ref]$d)) { return $d }
  return $null
}
function Test-RestEmpty($arr, [int]$r) {
  for ($c = 2; $c -le $arr.GetLength(1); $c++) {
    if ((Get-Str $arr $r $c).Trim() -ne '') { return $false }
  }
  return $true
}
# Titillína blokkar: "KG Bráða-..." (svið samtals) eða "KG03 Ferlideildir..." og ekkert annað í línunni
function Get-BlockKey($arr, [int]$r) {
  $a = (Get-Str $arr $r 1).Trim()
  if ($a -match '^KG(\d\d)?\s' -and (Test-RestEmpty $arr $r)) {
    if ($Matches[1]) { return 'KG' + $Matches[1] } else { return 'KG' }
  }
  return $null
}
# Dálkahausar: mánuðir 01–12 og "Samtals ÁÁÁÁ" (fyrst = yfirstandandi ár, svo fyrra ár)
function Read-Header($arr, [int]$r) {
  $h = @{ months = @{}; s26 = 0; s25 = 0; ar = 0; fyrraAr = 0 }
  for ($c = 2; $c -le $arr.GetLength(1); $c++) {
    $raw = $arr[$r, $c]
    $t = (Get-Str $arr $r $c).Trim()
    if ($t -match '^\d{1,2}$') {
      $m = [int]$t
      if ($m -ge 1 -and $m -le 12) { $h.months[$c] = $m - 1 }
    } elseif ($t -match '^Samtals\s+(\d{4})$') {
      if ($h.s26 -eq 0) { $h.s26 = $c; $h.ar = [int]$Matches[1] }
      elseif ($h.s25 -eq 0) { $h.s25 = $c; $h.fyrraAr = [int]$Matches[1] }
    }
  }
  return $h
}
function New-Line($arr, [int]$r, $h) {
  $mon = New-Object double[] 12
  foreach ($c in $h.months.Keys) {
    $v = Get-Num $arr $r $c
    if ($null -ne $v) { $mon[$h.months[$c]] = $v }
  }
  $s26 = Get-Num $arr $r $h.s26; if ($null -eq $s26) { $s26 = 0.0 }
  $s25 = Get-Num $arr $r $h.s25; if ($null -eq $s25) { $s25 = 0.0 }
  return @{ monthly = $mon; s26 = $s26; s25 = $s25 }
}
function Get-Block($store, [string]$key) {
  if (-not $store.Contains($key)) {
    $store[$key] = @{ exp = [ordered]@{}; vara = $null; svc = $null; atc = New-Object System.Collections.ArrayList; top = New-Object System.Collections.ArrayList; seen = @{} }
  }
  return $store[$key]
}

# ── Lestur blaða ───────────────────────────────────────────────────────────
# Gjaldabókun: bókhaldslyklar + samtölulínur vörukaupa og þjónustugjalda
function Parse-Gjald($arr, $store, $info) {
  $blk = $null; $h = $null
  for ($r = 1; $r -le $arr.GetLength(0); $r++) {
    $k = Get-BlockKey $arr $r
    if ($k) { $blk = Get-Block $store $k; $h = $null; continue }
    if (-not $blk) { continue }
    $a = (Get-Str $arr $r 1).Trim()
    if ($a -eq 'Gjaldategund bókhalds') {
      $h = Read-Header $arr $r
      if ($h.ar -and -not $info.ar) { $info.ar = $h.ar; $info.fyrraAr = $h.fyrraAr }
      continue
    }
    if (-not $h) { continue }
    # Oracle-útflutningur getur verið með tvíteknum röðum: fyrsta lína gildir
    if ($a -match '^(\d{6})\s') {
      $code = $Matches[1]
      if ($blk.seen.ContainsKey("G$code")) { continue }
      $blk.seen["G$code"] = 1
      $blk.exp[$code] = New-Line $arr $r $h
    } elseif ($a -like 'Lyf og önnur vara með álagi samtals*') {
      if (-not $blk.vara) { $blk.vara = New-Line $arr $r $h }
    } elseif ($a -like 'Þjónustugjöld samtals*') {
      if (-not $blk.svc) { $blk.svc = New-Line $arr $r $h }
    }
  }
}
# Yfirlit ATC flokka: ATC 1. stigs (A, B, C ...); "ATC 1 flokkun ekki til" og samtala ekki með
function Parse-Atc($arr, $store) {
  $blk = $null; $h = $null; $idx = 0
  for ($r = 1; $r -le $arr.GetLength(0); $r++) {
    $k = Get-BlockKey $arr $r
    if ($k) { $blk = Get-Block $store $k; $h = $null; continue }
    if (-not $blk) { continue }
    $a = (Get-Str $arr $r 1).Trim()
    if ($a -match '^ATC 1\s+lyfjaflokkar') { $h = Read-Header $arr $r; continue }
    if (-not $h) { continue }
    if ($a -match '^([A-Z])\s{2,}\S') {
      $code = $Matches[1]
      if ($blk.seen.ContainsKey("A$code")) { continue }
      $blk.seen["A$code"] = 1
      $ln = New-Line $arr $r $h
      $ln.code = $code; $ln.name = $a; $ln.idx = $idx++
      [void]$blk.atc.Add($ln)
    }
  }
}
# 30 kostnaðarmestu samheiti lyfja: allar línur með heiti og tölu 2026 (líka samtölulínur, eins og í frumriti)
function Parse-Top($arr, $store, $info) {
  $blk = $null; $c26 = 0; $c25 = 0; $idx = 0
  for ($r = 1; $r -le $arr.GetLength(0); $r++) {
    $a = (Get-Str $arr $r 1).Trim()
    if ($a -match '^Mánuðir til og með\s+(\d{1,2})') { $info.manTop = [int]$Matches[1]; continue }
    $k = Get-BlockKey $arr $r
    if ($k) { $blk = Get-Block $store $k; $c26 = 0; $c25 = 0; continue }
    if (-not $blk) { continue }
    if ($a -eq 'ATC 7 heiti') {
      for ($c = 2; $c -le $arr.GetLength(1); $c++) {
        $t = ((Get-Str $arr $r $c) -replace '\s+', ' ').Trim()
        if ($t -match '^Heildarverð í krónum (\d{4})$') {
          if ($c26 -eq 0) { $c26 = $c } elseif ($c25 -eq 0) { $c25 = $c }
        }
      }
      continue
    }
    if ($c26 -eq 0 -or $a -eq '') { continue }
    $s26 = Get-Num $arr $r $c26
    if ($null -eq $s26) { continue }
    if ($blk.seen.ContainsKey("T$a")) { continue }
    $blk.seen["T$a"] = 1
    $s25 = Get-Num $arr $r $c25; if ($null -eq $s25) { $s25 = 0.0 }
    [void]$blk.top.Add(@{ name = $a; s26 = $s26; s25 = $s25; idx = $idx++ })
  }
}

function Read-Workbook($xl, [string]$path) {
  $res = @{}
  $wb = $null
  try {
    $wb = $xl.Workbooks.Open($path, 0, $true)
    foreach ($ws in $wb.Worksheets) {
      $nm = $ws.Name
      $key = $null
      if ($nm -like 'Gjaldab*') { $key = 'gjald' }
      elseif ($nm -like 'Yfirlit ATC*') { $key = 'atc' }
      elseif ($nm -like '30 kostna*samheiti*') { $key = 'top' }
      if ($key) {
        $ur = $ws.UsedRange
        $lastR = $ur.Row + $ur.Rows.Count - 1
        $lastC = $ur.Column + $ur.Columns.Count - 1
        $rng = $ws.Range($ws.Cells(1, 1), $ws.Cells($lastR, $lastC))
        $v = $rng.Value2          # allt blaðið í einu lagi
        if ($v -isnot [array]) { $tmp = [object[,]]::new(1, 1); $v = $tmp }
        $res[$key] = $v
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($rng)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ur)
      }
      [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws)
    }
  } finally {
    if ($wb) { $wb.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
  }
  foreach ($k in 'gjald', 'atc', 'top') {
    if (-not $res.ContainsKey($k)) { throw "Blað fannst ekki ($k) í $path" }
  }
  return $res
}

function Find-File([string]$folder, [string]$prefix) {
  $f = @(Get-ChildItem -LiteralPath $folder -File | Where-Object {
      $_.Name -like "$prefix*" -and $_.Extension -match '^\.xls[xm]?$' -and $_.Name -notlike '~$*' })
  if ($f.Count -eq 0) { throw "Engin skrá sem byrjar á '$prefix' fannst í $folder" }
  if ($f.Count -gt 1) {
    $lyf = @($f | Where-Object { $_.Name -match 'Lyf' })
    if ($lyf.Count -ge 1) { $f = $lyf }
    $f = @($f | Sort-Object LastWriteTime -Descending)
    Write-Warning "Fleiri en ein '$prefix'-skrá; nota $($f[0].Name)"
  }
  return $f[0]
}

# ── Úttak ──────────────────────────────────────────────────────────────────
function ConvertTo-JsonText($o) {
  if ($null -eq $o) { return 'null' }
  if ($o -is [string]) {
    $s = $o.Replace('\', '\\').Replace('"', '\"').Replace("`r", '\r').Replace("`n", '\n').Replace("`t", '\t').Replace('</', '<\/')
    return '"' + $s + '"'
  }
  if ($o -is [bool]) { if ($o) { return 'true' } else { return 'false' } }
  if ($o -is [double] -or $o -is [int] -or $o -is [long] -or $o -is [decimal]) {
    $d = [double]$o
    if ($d -eq [Math]::Floor($d)) { return $d.ToString('0', $Inv) }
    return $d.ToString('R', $Inv)
  }
  if ($o -is [System.Collections.IDictionary]) {
    $parts = foreach ($k in $o.Keys) { (ConvertTo-JsonText ([string]$k)) + ': ' + (ConvertTo-JsonText $o[$k]) }
    return '{' + (@($parts) -join ', ') + '}'
  }
  if ($o -is [System.Collections.IEnumerable]) {
    $parts = foreach ($x in $o) { ConvertTo-JsonText $x }
    return '[' + (@($parts) -join ', ') + ']'
  }
  return (ConvertTo-JsonText ([string]$o))
}

function Rnd([double]$v) { return [Math]::Round($v) }
function Slice-Round($arr, [int]$n) { $o = New-Object double[] $n; for ($i = 0; $i -lt $n; $i++) { $o[$i] = Rnd $arr[$i] }; return , $o }

function Build-Gjald($blk, [int]$n) {
  $exp = [ordered]@{}
  foreach ($code in $blk.exp.Keys) {
    if (-not $EXP_DEF.Contains($code)) { continue }
    $e = $blk.exp[$code]
    $exp[$code] = [ordered]@{ name = $EXP_DEF[$code]; monthly = (Slice-Round $e.monthly $n); s26 = (Rnd $e.s26); s25 = (Rnd $e.s25) }
  }
  $zero = @{ monthly = (New-Object double[] 12); s26 = 0.0; s25 = 0.0 }
  $v = $blk.vara; if (-not $v) { $v = $zero }
  $s = $blk.svc;  if (-not $s) { $s = $zero }
  return [ordered]@{
    expenses = $exp
    vara_monthly = (Slice-Round $v.monthly $n); vara_s26 = (Rnd $v.s26); vara_s25 = (Rnd $v.s25)
    svc_monthly = (Slice-Round $s.monthly $n); svc_s26 = (Rnd $s.s26); svc_s25 = (Rnd $s.s25)
  }
}
function Build-Atc($blk, [int]$n) {
  $list = @($blk.atc | Sort-Object @{ Expression = { $_.s26 }; Descending = $true }, @{ Expression = { $_.idx }; Ascending = $true })
  $out = New-Object System.Collections.ArrayList
  foreach ($a in $list) {
    [void]$out.Add([ordered]@{ code = $a.code; name = $a.name; s26 = (Rnd $a.s26); s25 = (Rnd $a.s25); monthly = (Slice-Round $a.monthly $n) })
  }
  return , $out
}
function Build-Top($blk) {
  $list = @($blk.top | Sort-Object @{ Expression = { $_.s26 }; Descending = $true }, @{ Expression = { $_.idx }; Ascending = $true } | Select-Object -First 30)
  $out = New-Object System.Collections.ArrayList
  foreach ($t in $list) { [void]$out.Add([ordered]@{ name = $t.name; s26 = (Rnd $t.s26); s25 = (Rnd $t.s25) }) }
  return , $out
}

# ── Keyrsla ────────────────────────────────────────────────────────────────
if (-not (Test-Path -LiteralPath $Mappa)) { throw "Mappa fannst ekki: $Mappa" }
if ($Manudur -notmatch '^[\w\-]+$') { throw "Ógilt mánaðarheiti: $Manudur" }

$f81 = Find-File $Mappa '8.1'
$f82 = Find-File $Mappa '8.2'
Write-Host "8.1: $($f81.Name)"
Write-Host "8.2: $($f82.Name)"

$xl = $null
$wb81 = $null; $wb82 = $null
try {
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $xl.ScreenUpdating = $false
  $xl.EnableEvents = $false
  $xl.AskToUpdateLinks = $false
  $wb81 = Read-Workbook $xl $f81.FullName
  $wb82 = Read-Workbook $xl $f82.FullName
} finally {
  if ($xl) {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    $xl = $null
  }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

$svid = [ordered]@{}   # 8.1: eitt KG-samtölublokk
$eining = [ordered]@{} # 8.2: KG01–KG07
$info = @{ ar = 0; fyrraAr = 0; manTop = 0 }
$info82 = @{ ar = 0; fyrraAr = 0; manTop = 0 }
Parse-Gjald $wb81.gjald $svid $info
Parse-Atc   $wb81.atc   $svid
Parse-Top   $wb81.top   $svid $info
Parse-Gjald $wb82.gjald $eining $info82
Parse-Atc   $wb82.atc   $eining
Parse-Top   $wb82.top   $eining $info82

if (-not $svid.Contains('KG')) { throw "KG-samtölublokk fannst ekki í 8.1" }
$blend = $svid['KG']
if (-not $blend.vara) { throw "Samtölulína vörukaupa fannst ekki í 8.1" }

# Fjöldi mánaða
$n = $info.manTop
if (-not $n) {
  for ($i = 11; $i -ge 0; $i--) { if ($blend.vara.monthly[$i] -ne 0 -or ($blend.svc -and $blend.svc.monthly[$i] -ne 0)) { $n = $i + 1; break } }
  Write-Warning "Línan 'Mánuðir til og með' fannst ekki; ályktað út frá gögnum: $n mánuðir"
}
if ($info82.manTop -and $info82.manTop -ne $n) { Write-Warning "8.1 segir $n mánuði en 8.2 segir $($info82.manTop)" }
if ($info82.ar -and $info82.ar -ne $info.ar) { Write-Warning "Ár stemmir ekki milli 8.1 ($($info.ar)) og 8.2 ($($info82.ar))" }
$ar = $info.ar; $fyrra = $info.fyrraAr
if (-not $fyrra) { $fyrra = $ar - 1 }

# Afstemming: sviðseiningar samtals vs. svið
$sumKg = 0.0
foreach ($k in $eining.Keys) { if ($eining[$k].vara) { $sumKg += $eining[$k].vara.s26 } }
$mism = $blend.vara.s26 - $sumKg
Write-Host ("Afstemming vörukaup {0}: svið {1:N0} kr., sviðseiningar KG01–07 samtals {2:N0} kr., mismunur {3:N0} kr." -f $ar, $blend.vara.s26, $sumKg, $mism)
$vant = @($KG_DEF.Keys | Where-Object { -not $eining.Contains($_) })
if ($vant.Count) { Write-Warning "Sviðseiningar fundust ekki í 8.2: $($vant -join ', ')" }

$kgOut = [ordered]@{}
foreach ($k in $KG_DEF.Keys) {
  if (-not $eining.Contains($k)) { continue }
  $b = $eining[$k]
  $kgOut[$k] = [ordered]@{
    name = $KG_DEF[$k].name; color = $KG_DEF[$k].color
    gjald = (Build-Gjald $b $n); atc = (Build-Atc $b $n); top = (Build-Top $b)
  }
}

$meta = [ordered]@{
  manudur = $Manudur
  ar = $ar
  fyrraAr = $fyrra
  fjoldiManada = $n
  manudir = @($MAN_STUTT[0..($n - 1)])
  timabil = "Janúar–$($MAN_LANGT[$n - 1].ToLower()) $ar"
  timabilStutt = "Jan–$($MAN_BIL[$n - 1].ToLower())"
  heimildir = @($f81.Name, $f82.Name)
  utbuid = (Get-Date).ToString('yyyy-MM-dd HH:mm', $Inv)
}
$D = [ordered]@{
  meta = $meta
  blend = [ordered]@{ gjald = (Build-Gjald $blend $n); atc = (Build-Atc $blend $n); top = (Build-Top $blend) }
  kg = $kgOut
}

$outDir = Join-Path (Join-Path $Rot 'data') $Manudur
if (-not (Test-Path -LiteralPath $outDir)) { [void](New-Item -ItemType Directory -Path $outDir) }
$outFile = Join-Path $outDir 'lyf.js'
$txt = "// Lyfjakostnaður BLEND — búið til af scripts/lyf.ps1 úr: $($f81.Name) og $($f82.Name)`n" +
       'window.D = ' + (ConvertTo-JsonText $D) + ";`n"
[IO.File]::WriteAllText($outFile, $txt, (New-Object Text.UTF8Encoding($false)))
Write-Host ("Skrifað: {0}  ({1} mánuðir, {2})" -f $outFile, $n, $meta.timabil)
Write-Host ("Heildarlyfjakostnaður KG {0} (vörur + þjónusta): {1:N0} kr.; þar af 535100 Lyf: {2:N0} kr." -f $meta.timabilStutt, ($blend.vara.s26 + $(if ($blend.svc) { $blend.svc.s26 } else { 0 })), $(if ($blend.exp.Contains('535100')) { $blend.exp['535100'].s26 } else { 0 }))
