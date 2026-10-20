<#
.SYNOPSIS
  Býr til gagnaskrá fjárhagsmælaborðs BLEND (data\<Manudur>\fjarhagur.js) úr Heildaruppgjöri (Oracle/Hyperion).

.DESCRIPTION
  Les „Heildaruppgjör ….xlsx“ úr möppunni -Mappa gegnum Excel COM (read-only), finnur blaðið
  „Uppgjör Tegundir“ (eða fyrsta blað með dálkinum „Skipulag“), og reiknar fyrir KG (BLEND),
  sviðseiningar KG01–KG07 og öll viðföng þeirra: raun, áætlun, frávik og frávik % fyrir
  tegundirnar 0, 4, 5A, 5B, 50, 51, 510, 511 og 514. Upphæðir eru í m.kr. með einum aukastaf.

  Reglur:
   - Frávik = áætlun − raun. Neikvætt frávik = yfir áætlun (kostnaður).
   - Frávik % = frávik / |áætlun| × 100 (0 ef áætlun er 0).
   - Oracle-útflutningur er oft tvítekinn: aðeins fyrsta röð hverrar (Skipulag, Tegund) er notuð.
   - Viðfang (5 stafa kóði) tilheyrir síðustu sviðseiningu (KG0x) sem birtist á undan því í skránni.
   - 11131 og 11135 eru alltaf sameinuð: „11131+11135 Legudeild Minnisraskana / Líknardeild“.
   - Tímabil (fjöldi og heiti mánaða) ræðst af reitnum „Mánuður“ efst í Excel-skránni.

.EXAMPLE
  .\fjarhagur.ps1 -Mappa "C:\Users\arnija\OneDrive - Health\Claude\Gögn\Október" -Manudur '2026-09'
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Mappa,
  [Parameter(Mandatory = $true)][string]$Manudur
)

$ErrorActionPreference = 'Stop'
$inv = [Globalization.CultureInfo]::InvariantCulture
$Rot = Split-Path -Parent $PSScriptRoot          # rót appsins (Mælaborð BLEND)
$KODAR = @('0', '4', '5A', '5B', '50', '51', '510', '511', '514')
$SAMEINAD_HEITI = '11131+11135 Legudeild Minnisraskana / Líknardeild'
$SAMEINAD_KODAR = @('11131', '11135')

$MANUDIR_HEITI = @('janúar', 'febrúar', 'mars', 'apríl', 'maí', 'júní', 'júlí', 'ágúst', 'september', 'október', 'nóvember', 'desember')
$MANUDIR_STYTT = @{ 'jan' = 1; 'feb' = 2; 'mar' = 3; 'apr' = 4; 'maí' = 5; 'mai' = 5; 'jún' = 6; 'jun' = 6; 'júl' = 7; 'jul' = 7;
                    'ágú' = 8; 'agu' = 8; 'aug' = 8; 'sep' = 9; 'okt' = 10; 'oct' = 10; 'nóv' = 11; 'nov' = 11; 'des' = 12; 'dec' = 12 }

# ── Finna skrá ──────────────────────────────────────────────────────────────
function Finna-Skra {
  param([string]$Mappa, [string[]]$Mynstur, [string[]]$Utiloka)
  $skrar = @()
  foreach ($m in $Mynstur) { $skrar += @(Get-ChildItem -LiteralPath $Mappa -File -Filter $m -ErrorAction SilentlyContinue) }
  $skrar = @($skrar | Where-Object { $_.Name -notlike '~$*' } | Sort-Object FullName -Unique)
  foreach ($u in $Utiloka) { $skrar = @($skrar | Where-Object { $_.Name -notlike $u }) }
  if ($skrar.Count -eq 0) { return $null }
  # Nýjasta útgáfa: „uppfært“ fremst, síðan nýjasta breytingardagsetning.
  $skrar | Sort-Object @{ Expression = { if ($_.Name -match 'uppf') { 0 } else { 1 } } }, @{ Expression = { $_.LastWriteTime }; Descending = $true } |
    Select-Object -First 1
}

# Námundun að 1 aukastaf út frá nákvæmu tvíundargildi (G17), jafnt-á-jafnt eins og í upprunalegu vinnslunni.
function R1([double]$x) { [double][Math]::Round([decimal]::Parse($x.ToString('G17', $inv), [Globalization.NumberStyles]::Float, $inv), 1, [MidpointRounding]::ToEven) }
# Frávik % eins og í Hyperion: frávik / áætlun × 100; formerki snúið fyrir tekjur (4). 0 ef áætlun er 0.
function Fravik-Pct([string]$Kodi, [double]$F, [double]$A) {
  if ($A -eq 0) { return 0.0 }
  $p = $F / $A * 100; if ($Kodi -eq '4') { $p = -$p }
  R1 $p
}

# ── Lesa Excel (read-only, UsedRange.Value2 í heilu lagi) ───────────────────
function Lesa-Blad {
  param([string]$Slod, [string]$BladHeiti)
  $xl = $null; $wb = $null; $sheets = $null; $ws = $null; $ur = $null
  try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
    $wb = $xl.Workbooks.Open($Slod, 0, $true)
    $sheets = $wb.Worksheets
    $valid = $null
    for ($i = 1; $i -le $sheets.Count; $i++) {
      $s = $sheets.Item($i)
      if ($s.Name -eq $BladHeiti) { $valid = $s; break }
      [void][Runtime.InteropServices.Marshal]::ReleaseComObject($s)
    }
    if (-not $valid) { $valid = $sheets.Item(1); Write-Warning "Blaðið '$BladHeiti' fannst ekki — nota '$($valid.Name)'." }
    $ws = $valid
    $ur = $ws.UsedRange
    [pscustomobject]@{ Gildi = $ur.Value2; Blad = [string]$ws.Name }
  }
  finally {
    foreach ($o in @($ur, $ws, $sheets)) { if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
    if ($wb) { $wb.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
    if ($xl) { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
    $ur = $null; $ws = $null; $sheets = $null; $wb = $null; $xl = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}

function Tala($x) { if ($null -eq $x -or "$x" -eq '') { 0.0 } elseif ($x -is [double]) { $x } else { [double]::Parse("$x", $inv) } }

# ── JSON-ritari (heldur röð lykla, 2 bila inndráttur) ────────────────────────
function Json-Strengur([string]$s) {
  $sb = New-Object Text.StringBuilder
  [void]$sb.Append('"')
  foreach ($ch in $s.ToCharArray()) {
    switch ($ch) {
      '"' { [void]$sb.Append('\"') } '\' { [void]$sb.Append('\\') }
      "`n" { [void]$sb.Append('\n') } "`r" { [void]$sb.Append('\r') } "`t" { [void]$sb.Append('\t') }
      default { if ([int]$ch -lt 32) { [void]$sb.Append(('\u{0:x4}' -f [int]$ch)) } else { [void]$sb.Append($ch) } }
    }
  }
  [void]$sb.Append('"'); $sb.ToString()
}
function Til-Json($o, [int]$d = 0) {
  $p = '  ' * ($d + 1); $pe = '  ' * $d
  if ($null -eq $o) { return 'null' }
  if ($o -is [bool]) { if ($o) { return 'true' } else { return 'false' } }
  if ($o -is [int] -or $o -is [long]) { return $o.ToString($inv) }
  if ($o -is [double]) { return $o.ToString('0.0', $inv) }
  if ($o -is [string]) { return (Json-Strengur $o) }
  if ($o -is [System.Collections.IDictionary]) {
    $parts = foreach ($k in $o.Keys) { $p + (Json-Strengur ([string]$k)) + ': ' + (Til-Json $o[$k] ($d + 1)) }
    if (-not $parts) { return '{}' }
    return "{`n" + ($parts -join ",`n") + "`n$pe}"
  }
  if ($o -is [System.Collections.IEnumerable]) {
    $parts = foreach ($x in $o) { $p + (Til-Json $x ($d + 1)) }
    if (-not $parts) { return '[]' }
    return "[`n" + ($parts -join ",`n") + "`n$pe]"
  }
  return (Json-Strengur ([string]$o))
}

# ════════════════════════════════════════════════════════════════════════════
if (-not (Test-Path -LiteralPath $Mappa)) { throw "Mappan fannst ekki: $Mappa" }
if ($Manudur -notmatch '^[\w\-]+$') { throw "Ógilt -Manudur: $Manudur" }

$skra = Finna-Skra -Mappa $Mappa -Mynstur @('Heildaruppgj*.xlsx', '*Heildaruppgj*.xlsx', '*uppgj*r*.xlsx') -Utiloka @('*tkomusp*', '*Útkomusp*')
if (-not $skra) { throw "Engin Heildaruppgjörsskrá fannst í $Mappa" }
Write-Host "Les: $($skra.Name)"

$les = Lesa-Blad -Slod $skra.FullName -BladHeiti 'Uppgjör Tegundir'
$v = $les.Gildi
$nR = $v.GetLength(0); $nC = $v.GetLength(1)

# Hausröð og dálkar
$hRod = 0; $mTexti = $null
for ($r = 1; $r -le [Math]::Min(40, $nR); $r++) {
  if ("$($v[$r,1])".Trim() -eq 'Mánuður') { $mTexti = "$($v[$r,2])".Trim() }
  for ($c = 1; $c -le $nC; $c++) { if ("$($v[$r,$c])".Trim() -eq 'Skipulag') { $hRod = $r; break } }
  if ($hRod) { break }
}
if (-not $hRod) { throw "Fann ekki hausröð með 'Skipulag' í $($skra.Name)" }
$col = @{}
for ($c = 1; $c -le $nC; $c++) {
  $h = "$($v[$hRod,$c])".Trim()
  if ($h -eq 'Skipulag') { $col.S = $c }
  elseif ($h -eq 'Tegund') { $col.T = $c }
  elseif ($h -like 'Rauntölur*' -and -not $col.R) { $col.R = $c }
  elseif ($h -like 'Áætlun*' -and -not $col.A) { $col.A = $c }
  elseif ($h -like 'Frávik*' -and $h -notlike '*%*' -and -not $col.F) { $col.F = $c }
}
foreach ($k in 'S', 'T', 'R', 'A') { if (-not $col[$k]) { throw "Vantar dálk '$k' í hausröð ($($skra.Name))" } }

# Tímabil úr „Mánuður“ (t.d. „Sep 2026“)
if (-not $mTexti) { throw 'Fann ekki reitinn Mánuður efst í skránni.' }
if ($mTexti -notmatch '^\s*(\S+)\s+(\d{4})\s*$') { throw "Skil ekki Mánuður-reitinn: '$mTexti'" }
$mHluti = $Matches[1].ToLower(); $ar = [int]$Matches[2]
$nMan = $null
$ix = [Array]::IndexOf($MANUDIR_HEITI, $mHluti)
if ($ix -ge 0) { $nMan = $ix + 1 } else { $k3 = $mHluti.Substring(0, [Math]::Min(3, $mHluti.Length)); if ($MANUDIR_STYTT.ContainsKey($k3)) { $nMan = $MANUDIR_STYTT[$k3] } }
if (-not $nMan) { throw "Óþekktur mánuður: '$mTexti'" }
$manudir = @($MANUDIR_HEITI[0..($nMan - 1)])
$sidasti = $manudir[-1]
$Cap = { param($s) $s.Substring(0, 1).ToUpper() + $s.Substring(1) }
$timabil = if ($nMan -eq 1) { 'janúar' } else { "janúar–$sidasti" }
$timabilLangt = if ($nMan -eq 1) { "Janúar $ar" } else { "Janúar–$sidasti $ar" }
$timabilTexti = if ($nMan -eq 1) { "janúar $ar" } else { "janúar til $sidasti $ar" }

# ── Safna röðum (fyrsta tilvik hverrar Skipulag|Tegund), með stigveldi ───────
$rec = @{}
$born = [ordered]@{}               # KG0x -> lista af Skipulag-strengjum (í röð)
$foreldri = $null; $kgTopp = $null; $tvitekid = 0
for ($r = $hRod + 1; $r -le $nR; $r++) {
  $S = "$($v[$r,$col.S])".Trim()
  if (-not $S) { continue }
  if ($S -match '^([A-Z]{2}\d{0,2})\s+-\s') {
    $org = $Matches[1]
    if ($org -match '^KG0\d$') { $foreldri = $org; if (-not $born.Contains($org)) { $born[$org] = New-Object System.Collections.ArrayList } }
    else { $foreldri = $null; if ($org -eq 'KG' -and -not $kgTopp) { $kgTopp = $S } }
  }
  elseif ($S -match '^\d{5}\s+-\s' -and $foreldri) {
    if (-not $born[$foreldri].Contains($S)) { [void]$born[$foreldri].Add($S) }
  }
  $T = "$($v[$r,$col.T])".Trim()
  if ($T -notmatch '^(\w+)\s+-\s') { continue }
  $kodi = $Matches[1]
  if ($KODAR -notcontains $kodi) { continue }
  $lykill = "$S|$kodi"
  if ($rec.ContainsKey($lykill)) { $tvitekid++; continue }
  $raun = Tala $v[$r,$col.R]; $aetl = Tala $v[$r,$col.A]
  $frav = if ($col.F) { Tala $v[$r,$col.F] } else { $aetl - $raun }
  $rec[$lykill] = @($raun, $aetl, $frav)
}
if (-not $kgTopp) { throw 'Fann ekki KG-línu (BLEND samtals) í skránni.' }
$v = $null

# KG0x-heiti eins og þau birtast í skránni
$kgHeiti = @{}
foreach ($key in $rec.Keys) { $s = $key.Split('|')[0]; if ($s -match '^(KG0\d)\s+-\s') { $kgHeiti[$Matches[1]] = $s } }

function Byggja-Hnut([string]$Skipulag) {
  $n = [ordered]@{}
  foreach ($k in $KODAR) {
    $x = $rec["$Skipulag|$k"]
    if ($x) { $r = $x[0]; $a = $x[1]; $f = $x[2] } else { $r = 0.0; $a = 0.0; $f = 0.0 }
    $n[$k] = [ordered]@{ raun = (R1 ($r / 1e6)); aetlun = (R1 ($a / 1e6)); fravik = (R1 ($f / 1e6)); fravik_pct = (Fravik-Pct $k $f $a) }
  }
  $n
}
# Sameinað viðfang (11131+11135): summa námundaðra m.kr.-gilda hvers hluta; % reiknað af summunum.
function Sameina-Hnuta($Hnutar) {
  $n = [ordered]@{}
  foreach ($k in $KODAR) {
    $r = 0.0; $a = 0.0; $f = 0.0
    foreach ($h in $Hnutar) { $r += $h[$k].raun; $a += $h[$k].aetlun; $f += $h[$k].fravik }
    $n[$k] = [ordered]@{ raun = (R1 $r); aetlun = (R1 $a); fravik = (R1 $f); fravik_pct = (Fravik-Pct $k $f $a) }
  }
  $n
}

$svid = [ordered]@{}; $vidfong = [ordered]@{}
foreach ($kg in @($born.Keys | Sort-Object)) {
  $svid[$kg] = Byggja-Hnut $kgHeiti[$kg]
  $listi = New-Object System.Collections.ArrayList
  $sameinad = @($born[$kg] | Where-Object { $SAMEINAD_KODAR -contains $_.Substring(0, 5) })
  $sameinadSkrifad = $false
  foreach ($s in $born[$kg]) {
    if ($SAMEINAD_KODAR -contains $s.Substring(0, 5)) {
      if ($sameinadSkrifad) { continue }
      $n = Sameina-Hnuta @($sameinad | ForEach-Object { Byggja-Hnut $_ }); $n['label'] = $SAMEINAD_HEITI; $sameinadSkrifad = $true
    }
    else { $n = Byggja-Hnut $s; $n['label'] = $s }
    $n['kg'] = $kg
    [void]$listi.Add($n)
  }
  $vidfong[$kg] = $listi
}

# ── Sannprófun: summa sviðseininga = BLEND, summa viðfanga = sviðseining ─────
$blend = Byggja-Hnut $kgTopp
$sumKg = 0.0; foreach ($kg in $svid.Keys) { $x = $rec["$($kgHeiti[$kg])|0"]; if ($x) { $sumKg += $x[0] } }
$kgRaun = $rec["$kgTopp|0"][0]
if ([Math]::Abs($sumKg - $kgRaun) -gt 1e5) { Write-Warning ("Summa KG01–KG07 ({0:N0}) víkur frá KG samtals ({1:N0})." -f $sumKg, $kgRaun) }
foreach ($kg in $svid.Keys) {
  $sv = 0.0; foreach ($s in $born[$kg]) { $x = $rec["$s|0"]; if ($x) { $sv += $x[0] } }
  $x = $rec["$($kgHeiti[$kg])|0"]
  if ($x -and [Math]::Abs($sv - $x[0]) -gt 1e5) { Write-Warning ("{0}: summa viðfanga ({1:N0}) víkur frá sviðseiningu ({2:N0})." -f $kg, $sv, $x[0]) }
}

# ── Handskrifaðar skýringar (ekki hægt að leiða af Excel) ────────────────────
# Afritaðar úr júlí-mælaborðinu og merktar til yfirferðar (gagnagaedi_yfirfara = true).
$gagnagaedi = @(
  [ordered]@{ titill = 'Mars-tímar'; texti = 'Gögn um yfirvinnu í mars eru óáreiðanleg vegna tæknilegs vandamáls í VS-gagnagrunni. Eingöngu janúar og febrúar eru notuð til kostnaðarmats þar sem við á.' },
  [ordered]@{ titill = '09843 Háþrýsti'; texti = 'Þjónustusamningur við þriðja aðila rann út í ársbyrjun 2026. Kostnaður sem áður féll á þennan samning er nú á viðfanginu en fjárheimild var ekki uppfærð að sama skapi.' },
  [ordered]@{ titill = '09083 Námslæknar'; texti = 'Mánaðarlaun námslækna eru greidd af skrifstofu lækninga en yfirvinnu- og vaktavinnukostnaður þeirra er borinn af deildunum. Þessar tölur eru því ekki samanburðarhæfar við önnur viðföng.' }
)

$D = [ordered]@{
  meta = [ordered]@{
    manudur = $Manudur; ar = $ar; fjoldi_manada = $nMan; manudir = $manudir
    fyrsti_manudur = $manudir[0]; sidasti_manudur = $sidasti
    timabil = $timabil; timabil_langt = $timabilLangt; timabil_texti = $timabilTexti
    heimild = $skra.Name; blad = $les.Blad; eining = 'm.kr.'
    buid_til = (Get-Date).ToString('yyyy-MM-dd HH:mm', $inv)
  }
  texti = [ordered]@{ gagnagaedi_yfirfara = $true; gagnagaedi = $gagnagaedi }
  blend = $blend
  svidseining = $svid
  vidfong = $vidfong
}

$utMappa = Join-Path $Rot "data\$Manudur"
New-Item -ItemType Directory -Force -Path $utMappa | Out-Null
$utSkra = Join-Path $utMappa 'fjarhagur.js'
$js = "// Fjárhagsmælaborð BLEND — gögn fyrir $timabilTexti.`n" +
      "// Búið til sjálfvirkt af scripts/fjarhagur.ps1 úr '$($skra.Name)'. Ekki breyta handvirkt.`n" +
      "window.D = " + (Til-Json $D 0) + ";`n"
[IO.File]::WriteAllText($utSkra, $js, (New-Object Text.UTF8Encoding($false)))

$b0 = $blend['0']
Write-Host ("Tímabil: {0} ({1} mán.)  Tvíteknar raðir hunsaðar: {2}" -f $timabilLangt, $nMan, $tvitekid)
Write-Host ("BLEND rekstur: raun {0} / áætlun {1} / frávik {2} m.kr. ({3}%)" -f $b0.raun, $b0.aetlun, $b0.fravik, $b0.fravik_pct)
Write-Host ("Sviðseiningar: {0}  Viðföng: {1}" -f $svid.Count, (($vidfong.Values | ForEach-Object { $_.Count }) | Measure-Object -Sum).Sum)
Write-Host "Skrifað: $utSkra"
$mjs = Join-Path $Rot 'data\manudir.js'
if ((Test-Path $mjs) -and -not ((Get-Content -Raw -Encoding UTF8 $mjs) -match [regex]::Escape("'$Manudur'"))) {
  Write-Warning "Mánuðurinn '$Manudur' er ekki í data\manudir.js — bættu honum við svo hann birtist í appinu."
}
