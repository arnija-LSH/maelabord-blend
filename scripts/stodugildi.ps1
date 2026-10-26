<#
  stodugildi.ps1 — Býr til gagnaskrá fyrir mælaborðið „Stöðugildi“ (key = stodugildi).

  Notkun:
    powershell -ExecutionPolicy Bypass -File stodugildi.ps1 -Mappa "C:\...\Gögn\Október" -Manudur 2026-09

  Les (read-only, gegnum Excel COM):
    1. Yfirlit dagvinnustöðugilda (*dagvinnust*KG*.xlsx)  – blöð „Dv.stg. - sviðsein.“ og „Dv.stg. - viðföng“
    2. 5.1 Yfirvinnugreining (*Yfirvinnugreining*.xlsx)    – blöð „Yfirvinnugreining -Svið“, „- sviðsein.“, „- viðföng“
    3. Veikindagreining (Veikindagreining*.xlsx)           – blöð „Tímabilið“ og „Skipulagseiningar“
  Skrifar:  <rót>\data\<Manudur>\stodugildi.js   (window.D = {...};)

  Reglur:
    - Fjöldi og heiti mánaða ráðast af mánaðardálkum (01..12) í dagvinnustöðugildaskránni.
    - 11131 og 11135 eru alltaf sameinaðar í „11131+11135 Legudeild Minnisraskana / Líknardeild“.
    - Frávik = raun − áætlun (eins og í Excel-skránni): jákvætt = fleiri stöðugildi en áætlun.
    - Engin nöfn einstaklinga eru lesin; aðeins samantektarblöð.
#>
param(
  [Parameter(Mandatory = $true)][string]$Mappa,
  [Parameter(Mandatory = $true)][string]$Manudur
)
$ErrorActionPreference = 'Stop'
$Inv = [Globalization.CultureInfo]::InvariantCulture
$Rot = Split-Path -Parent $PSScriptRoot
$KEY = 'stodugildi'

# ── Fastar (ekki í mánaðargögnum) ──────────────────────────────────────────────
$KG_LITIR = [ordered]@{ KG01 = '#1F4E79'; KG02 = '#117A65'; KG03 = '#7D3C98'; KG04 = '#B7950B'; KG05 = '#1A5276'; KG06 = '#784212'; KG07 = '#616A6B' }
$MAN_LYKLAR = @('jan','feb','mar','apr','mai','jun','jul','agu','sep','okt','nov','des')
$MAN_STUTT  = @('Jan','Feb','Mar','Apr','Maí','Jún','Júl','Ágú','Sep','Okt','Nóv','Des')
$MAN_FULLT  = @('janúar','febrúar','mars','apríl','maí','júní','júlí','ágúst','september','október','nóvember','desember')
$SAMEINA_KODAR = @('11131','11135')
$SAMEINA_KODI  = '11131+11135'
$SAMEINA_NAFN  = '11131+11135 Legudeild Minnisraskana / Líknardeild'
$SAMEINA_VK    = @('^Legudeild minnisraskana', '^Líknardeild aldraðra')
$SAMEINA_VK_NAFN = 'Legudeild Minnisraskana / Líknardeild'

$Vidvaranir = New-Object System.Collections.ArrayList
function Vidvorun([string]$t) { [void]$Vidvaranir.Add($t); Write-Warning $t }

# ── Hjálparföll ───────────────────────────────────────────────────────────────
function Num($v) {
  if ($null -eq $v) { return $null }
  if ($v -is [double] -or $v -is [int] -or $v -is [decimal]) { return [double]$v }
  $s = "$v".Trim()
  if ($s -eq '') { return $null }
  $d = 0.0
  if ([double]::TryParse($s, [Globalization.NumberStyles]::Float, $Inv, [ref]$d)) { return $d }
  return $null
}
function N0($v) { $n = Num $v; if ($null -eq $n) { return 0.0 } return $n }
# Slétting: námundað út frá nákvæmu tvíundargildi talsins með „round half to even“
# (sama regla og round() í Python, sem frumritið var reiknað með).
Add-Type -AssemblyName System.Numerics
function Rnd([double]$x, [int]$d) {
  if ($x -eq 0 -or [double]::IsNaN($x) -or [double]::IsInfinity($x)) { return 0.0 }
  $neg = $x -lt 0; $a = [Math]::Abs($x)
  $bits = [BitConverter]::DoubleToInt64Bits($a)
  $e = [int](($bits -shr 52) -band 0x7FF); $m = $bits -band 0xFFFFFFFFFFFFF
  if ($e -eq 0) { $e = 1 } else { $m = $m + 4503599627370496 }
  $e = $e - 1075
  $num = [System.Numerics.BigInteger]::Multiply([System.Numerics.BigInteger]$m, [System.Numerics.BigInteger]::Pow(10, $d))
  if ($e -ge 0) { return $x }
  $den = [System.Numerics.BigInteger]::Pow(2, -$e)
  $rem = [System.Numerics.BigInteger]::Zero
  $q = [System.Numerics.BigInteger]::DivRem($num, $den, [ref]$rem)
  $c = [System.Numerics.BigInteger]::Compare([System.Numerics.BigInteger]::Multiply($rem, 2), $den)
  if ($c -gt 0 -or ($c -eq 0 -and -not $q.IsEven)) { $q = $q + 1 }
  $r = [double]$q / [Math]::Pow(10, $d)
  if ($neg) { $r = -$r }
  if ($r -eq 0) { $r = 0.0 }
  return $r
}
function Txt($v) { if ($null -eq $v) { return '' } return ("$v").Trim() }
function Cell($a, [int]$r, [int]$c) {
  if ($r -lt 1 -or $c -lt 1 -or $r -gt $a.GetLength(0) -or $c -gt $a.GetLength(1)) { return $null }
  return $a[$r, $c]
}
function NormNafn([string]$s) { return ($s -replace '\s', '').ToLowerInvariant() }

function Finna-Skra([string]$lysing, [scriptblock]$sia) {
  $c = @(Get-ChildItem -LiteralPath $Mappa -File | Where-Object { $_.Name -notlike '~$*' -and $_.Extension -match '^\.xls[xm]$' } | Where-Object $sia)
  return $c
}

# ── Lesa Excel ────────────────────────────────────────────────────────────────
function Lesa-Blod($xl, [string]$slod, [string[]]$blod) {
  $ut = @{}
  $wb = $null
  try {
    $wb = $xl.Workbooks.Open($slod, 0, $true)
    $til = @{}
    foreach ($ws in $wb.Worksheets) { $til[(NormNafn $ws.Name)] = $ws.Name; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws) }
    foreach ($b in $blod) {
      $nn = NormNafn $b
      if (-not $til.ContainsKey($nn)) { $ut[$b] = $null; continue }
      $ws = $wb.Worksheets.Item($til[$nn])
      $ur = $ws.UsedRange
      # Value2 í heilu lagi; endurstillt þannig að (1,1) = efsta vinstra horn A1 svo dálkanúmer séu raunveruleg
      $v = $ur.Value2
      $r0 = $ur.Row; $c0 = $ur.Column
      $nr = $v.GetLength(0) + $r0 - 1; $nc = $v.GetLength(1) + $c0 - 1
      $a = New-Object 'object[,]' ($nr + 1), ($nc + 1)
      $lr = $v.GetLowerBound(0); $lc = $v.GetLowerBound(1)
      for ($i = 0; $i -lt $v.GetLength(0); $i++) { for ($j = 0; $j -lt $v.GetLength(1); $j++) { $a[($i + $r0), ($j + $c0)] = $v[($i + $lr), ($j + $lc)] } }
      $ut[$b] = $a
      [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ur)
      [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws)
    }
  } finally {
    if ($wb) { $wb.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
  }
  return $ut
}
# Fylkin eru (nr+1)x(nc+1) með ónotaðri línu/dálki 0; Rows/Cols skila síðasta nothæfa vísi.
function Rows($a) { return $a.GetLength(0) - 1 }
function Cols($a) { return $a.GetLength(1) - 1 }

# ── Finna skrár ───────────────────────────────────────────────────────────────
if (-not (Test-Path -LiteralPath $Mappa)) { throw "Mappa finnst ekki: $Mappa" }
$dvKand = Finna-Skra 'dagvinnustöðugildi' { $_.Name -match 'dagvinnust' }
$yvKand = Finna-Skra 'yfirvinnugreining' { $_.Name -match 'Yfirvinnugreining' }
$vkKand = Finna-Skra 'veikindagreining' { $_.Name -match '^Veikindagreining' }
if ($dvKand.Count -eq 0) { throw "Fann enga skrá «Yfirlit dagvinnustöðugilda…» í $Mappa" }
if ($yvKand.Count -eq 0) { throw "Fann enga skrá «…Yfirvinnugreining…» í $Mappa" }
if ($vkKand.Count -eq 0) { throw "Fann enga skrá «Veikindagreining…» í $Mappa" }
# Raða: skrár með „KG“ í heiti fyrst, svo nýjasta
$dvKand = @($dvKand | Sort-Object @{ e = { if ($_.Name -match '\bKG\b') { 0 } else { 1 } } }, @{ e = { $_.LastWriteTime }; Descending = $true })
$yvKand = @($yvKand | Sort-Object @{ e = { if ($_.Name -match '\bKG\b') { 0 } else { 1 } } }, @{ e = { $_.LastWriteTime }; Descending = $true })
$vkKand = @($vkKand | Sort-Object @{ e = { if ($_.Name -match '\bKG\b') { 0 } else { 1 } } }, @{ e = { $_.LastWriteTime }; Descending = $true })

$xl = $null
try {
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
  try { $xl.AskToUpdateLinks = $false } catch {}

  # Dagvinnustöðugildi: velja skrá sem nær yfir allt sviðið (titill í „Dv.stg. - sviðsein.“ byrjar á „KG “),
  # t.d. ekki „Lungnalækningar - Yfirlit dagvinnustöðugilda.xlsx“ sem nær aðeins yfir eitt viðfang.
  $dvSkra = $null; $DV = $null
  foreach ($k in $dvKand) {
    $b = Lesa-Blod $xl $k.FullName @('Dv.stg. - sviðsein.', 'Dv.stg. - viðföng')
    if ($null -eq $b['Dv.stg. - sviðsein.'] -or $null -eq $b['Dv.stg. - viðföng']) { continue }
    $a = $b['Dv.stg. - sviðsein.']; $titill = ''
    for ($r = 1; $r -le [Math]::Min(3, (Rows $a)); $r++) { for ($c = 1; $c -le (Cols $a); $c++) { $t = Txt (Cell $a $r $c); if ($t) { $titill = $t; break } }; if ($titill) { break } }
    if ($titill -match '^KG ') { $dvSkra = $k; $DV = $b; break }
    else { Write-Host "Sleppi $($k.Name) (titill: $titill)" }
  }
  if (-not $dvSkra) { throw "Engin dagvinnustöðugildaskrá fyrir allt KG-sviðið fannst í $Mappa" }
  $yvSkra = $yvKand[0]
  $YV = Lesa-Blod $xl $yvSkra.FullName @('Yfirvinnugreining -Svið', 'Yfirvinnugreining - sviðsein.', 'Yfirvinnugreining - viðföng')
  $vkSkra = $vkKand[0]
  $VK = Lesa-Blod $xl $vkSkra.FullName @('Tímabilið', 'Skipulagseiningar')
} finally {
  if ($xl) { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
foreach ($p in @(@($YV, 'Yfirvinnugreining -Svið'), @($YV, 'Yfirvinnugreining - sviðsein.'), @($YV, 'Yfirvinnugreining - viðföng'), @($VK, 'Tímabilið'), @($VK, 'Skipulagseiningar'))) {
  if ($null -eq $p[0][$p[1]]) { throw "Blaðið «$($p[1])» fannst ekki." }
}
Write-Host "Dagvinnustöðugildi: $($dvSkra.Name)"
Write-Host "Yfirvinna:          $($yvSkra.Name)"
Write-Host "Veikindi:           $($vkSkra.Name)"

# ── Dálkar í dagvinnustöðugildablöðum ─────────────────────────────────────────
function Haus-Dalkar($a, [int]$r) {
  $h = [ordered]@{ manudir = [ordered]@{}; ar = $null; medalt = $null; medalt_fyrra = $null; mism = $null; mism_pct = $null; aatl = $null; frav = $null; frav_pct = $null }
  for ($c = 1; $c -le (Cols $a); $c++) {
    $t = Txt (Cell $a $r $c)
    if ($t -match '^(0[1-9]|1[0-2])$') { $h.manudir[$t] = $c }
    elseif ($t -match '^Meðaltal (\d{4})$') { if ($null -eq $h.medalt) { $h.medalt = $c; $h.ar = [int]$Matches[1] } else { $h.medalt_fyrra = $c } }
    elseif ($t -eq 'Mismunur ára') { $h.mism = $c }
    elseif ($t -eq 'Mismunur ára %') { $h.mism_pct = $c }
    elseif ($t -match '^Áætlun \d{4}$') { $h.aatl = $c }
    elseif ($t -eq 'Frávik frá áætlun') { $h.frav = $c }
    elseif ($t -eq 'Frávik frá áætlun %') { $h.frav_pct = $c }
  }
  return $h
}
function Er-Haus($a, [int]$r) {
  for ($c = 1; $c -le (Cols $a); $c++) { if ((Txt (Cell $a $r $c)) -match '^Meðaltal \d{4}$') { return $true } }
  return $false
}
function Fyrsti-Textadalkur($a) {
  for ($c = 1; $c -le (Cols $a); $c++) { for ($r = 1; $r -le [Math]::Min(10, (Rows $a)); $r++) { if (Txt (Cell $a $r $c)) { return $c } } }
  return 1
}

# Les línu: skilar hráum tölum (óslétt)
function Lesa-Lina($a, [int]$r, $h) {
  $o = [ordered]@{ man = [ordered]@{}; medalt = 0.0; medalt_fyrra = 0.0; mism = 0.0; mism_pct = 0.0; aatl = 0.0; frav = 0.0; frav_pct = 0.0; hefur_tolur = $false }
  foreach ($m in $h.manudir.Keys) { $n = Num (Cell $a $r $h.manudir[$m]); if ($null -ne $n) { $o.hefur_tolur = $true }; $o.man[$m] = $(if ($null -eq $n) { 0.0 } else { $n }) }
  foreach ($f in 'medalt', 'medalt_fyrra', 'mism', 'mism_pct', 'aatl', 'frav', 'frav_pct') {
    if ($h[$f]) { $n = Num (Cell $a $r $h[$f]); if ($null -ne $n) { $o[$f] = $n; $o.hefur_tolur = $true } }
  }
  return $o
}

# ── KG (Dv.stg. - sviðsein.) ──────────────────────────────────────────────────
$A = $DV['Dv.stg. - sviðsein.']
$lc = Fyrsti-Textadalkur $A   # dálkur með heiti sviðseiningar (B)
$H = $null
$kgHrar = [ordered]@{}
for ($r = 1; $r -le (Rows $A); $r++) {
  if (Er-Haus $A $r) { $H = Haus-Dalkar $A $r; continue }
  if (-not $H) { continue }
  $lab = Txt (Cell $A $r $lc); $stett = Txt (Cell $A $r ($lc + 1))
  if ($lab -match '^(KG\d\d) (.+)$' -and $stett -eq '') {
    $kode = $Matches[1]
    $l = Lesa-Lina $A $r $H
    if (-not $l.hefur_tolur) { continue }
    if ($kgHrar.Contains($kode)) { Vidvorun "Tvítekin samtalslína fyrir $kode í «Dv.stg. - sviðsein.» — fyrsta línan notuð."; continue }
    $l['name'] = $lab; $l['label'] = $Matches[2]
    $kgHrar[$kode] = $l
  }
}
if ($kgHrar.Count -eq 0) { throw 'Engar KG-samtalslínur fundust í „Dv.stg. - sviðsein.“' }
$manKodar = @($H.manudir.Keys)           # t.d. 01..09
$AR = $H.ar
if (-not $AR) { throw 'Fann ekki ár (Meðaltal ÁÁÁÁ) í haus.' }
$manLyklar = @($manKodar | ForEach-Object { $MAN_LYKLAR[[int]$_ - 1] })
$manStutt  = @($manKodar | ForEach-Object { $MAN_STUTT[[int]$_ - 1] })
$sidastiIdx = [int]$manKodar[-1] - 1
$fyrstiIdx = [int]$manKodar[0] - 1

function Til-Hlut($l, [string[]]$forskeyti) {
  $o = [ordered]@{}
  for ($i = 0; $i -lt $manKodar.Count; $i++) { $o[$manLyklar[$i]] = Rnd $l.man[$manKodar[$i]] 2 }
  $o['medalt26'] = Rnd $l.medalt 2
  $o['medalt25'] = Rnd $l.medalt_fyrra 2
  $o['mism'] = Rnd $l.mism 2
  $o['mism_pct'] = Rnd $l.mism_pct 4
  $o['aatl'] = Rnd $l.aatl 2
  $o['frav'] = Rnd $l.frav 2
  $o['frav_pct'] = Rnd $l.frav_pct 4
  return $o
}

$kg_totals = [ordered]@{}
foreach ($kode in $kgHrar.Keys) {
  $l = $kgHrar[$kode]
  $o = [ordered]@{ name = $l.name; short = $l.label }
  $t = Til-Hlut $l
  foreach ($k in $t.Keys) { $o[$k] = $t[$k] }
  $o['label'] = $l.label; $o['code'] = $kode
  $o['color'] = $(if ($KG_LITIR.Contains($kode)) { $KG_LITIR[$kode] } else { Vidvorun "Enginn litur skilgreindur fyrir $kode — grár notaður."; '#888888' })
  $kg_totals[$kode] = $o
}

# ── Viðföng (Dv.stg. - viðföng) ───────────────────────────────────────────────
$A = $DV['Dv.stg. - viðföng']
$lc = Fyrsti-Textadalkur $A
$H = $null; $kgNu = $null
$vfHrar = New-Object System.Collections.ArrayList
$vfSed = @{}
# Stéttarfélagslínur hvers viðfangs: fyrsta línan ber heiti viðfangs, þær næstu hafa autt heiti
$stHrar = @{}; $stNu = $null
for ($r = 1; $r -le (Rows $A); $r++) {
  if (Er-Haus $A $r) { $H = Haus-Dalkar $A $r; continue }
  $lab = Txt (Cell $A $r $lc); $stett = Txt (Cell $A $r ($lc + 1))
  if ($lab -match '^(KG\d\d) ' -and $stett -eq '') { $kgNu = $Matches[1]; continue }
  if (-not $H) { continue }
  if ($stett -ne '' -and $stett -ne 'Stöðugildi:') {
    if ($lab -match '^(\d{5}) ') { $stNu = $(if ($vfSed.ContainsKey($Matches[1])) { $null } else { $Matches[1] }) }
    elseif ($lab -ne '') { $stNu = $null }
    if (-not $stNu) { continue }
    $l = Lesa-Lina $A $r $H
    if (-not $l.hefur_tolur) { continue }
    $l['name'] = $stett
    if (-not $stHrar.ContainsKey($stNu)) { $stHrar[$stNu] = New-Object System.Collections.ArrayList }
    [void]$stHrar[$stNu].Add($l)
    continue
  }
  if ($lab -match '^(\d{5}) ' -and $stett -eq '') {
    $kode = $Matches[1]
    $l = Lesa-Lina $A $r $H
    # Viðfang tekið með ef það hefur raun 2026 eða áætlun 2026 (viðföng með aðeins 2025-tölur sleppt)
    $hefurRaun = $false; foreach ($mk in $l.man.Keys) { if ($l.man[$mk] -ne 0) { $hefurRaun = $true } }
    if (-not $hefurRaun -and $l.medalt -eq 0 -and $l.aatl -eq 0) { continue }
    if ($vfSed.ContainsKey($kode)) { Vidvorun "Tvítekin samtalslína fyrir viðfang $kode — fyrsta línan notuð."; continue }
    $vfSed[$kode] = $true
    $l['kg'] = $kgNu; $l['code'] = $kode; $l['name'] = $lab
    [void]$vfHrar.Add($l)
  }
}

# Sameina 11131 + 11135
$sam = @($vfHrar | Where-Object { $SAMEINA_KODAR -contains $_.code })
if ($sam.Count -gt 0) {
  $m = [ordered]@{ man = [ordered]@{}; medalt = 0.0; medalt_fyrra = 0.0; aatl = 0.0; frav = 0.0; frav_pct = 0.0; kg = $sam[0].kg; code = $SAMEINA_KODI; name = $SAMEINA_NAFN }
  foreach ($mk in $manKodar) { $m.man[$mk] = 0.0 }
  foreach ($s in $sam) {
    foreach ($mk in $manKodar) { $m.man[$mk] += $s.man[$mk] }
    $m.medalt += $s.medalt; $m.medalt_fyrra += $s.medalt_fyrra; $m.aatl += $s.aatl; $m.frav += $s.frav
    if ($s.kg -ne $m.kg) { Vidvorun "11131 og 11135 eru í mismunandi sviðseiningum ($($sam[0].kg) / $($s.kg))." }
  }
  $m.frav_pct = $(if ($m.aatl -ne 0) { $m.frav / $m.aatl } else { 0.0 })
  $m.mism = $m.medalt - $m.medalt_fyrra
  $m.mism_pct = $(if ($m.medalt_fyrra -ne 0) { $m.mism / $m.medalt_fyrra } else { 0.0 })
  # Stéttarfélög beggja viðfanga lögð saman eftir heiti
  $stSam = [ordered]@{}
  foreach ($kd in $SAMEINA_KODAR) {
    if (-not $stHrar.ContainsKey($kd)) { continue }
    foreach ($s in $stHrar[$kd]) {
      if (-not $stSam.Contains($s.name)) {
        $n = [ordered]@{ man = [ordered]@{}; medalt = 0.0; medalt_fyrra = 0.0; aatl = 0.0; frav = 0.0; name = $s.name }
        foreach ($mk in $manKodar) { $n.man[$mk] = 0.0 }
        $stSam[$s.name] = $n
      }
      $n = $stSam[$s.name]
      foreach ($mk in $manKodar) { $n.man[$mk] += $s.man[$mk] }
      $n.medalt += $s.medalt; $n.medalt_fyrra += $s.medalt_fyrra; $n.aatl += $s.aatl; $n.frav += $s.frav
    }
  }
  $stHrar[$SAMEINA_KODI] = New-Object System.Collections.ArrayList
  foreach ($n in $stSam.Values) {
    $n['mism'] = $n.medalt - $n.medalt_fyrra
    $n['mism_pct'] = $(if ($n.medalt_fyrra -ne 0) { $n.mism / $n.medalt_fyrra } else { 0.0 })
    $n['frav_pct'] = $(if ($n.aatl -ne 0) { $n.frav / $n.aatl } else { 0.0 })
    [void]$stHrar[$SAMEINA_KODI].Add($n)
  }
  $nyr = New-Object System.Collections.ArrayList
  $sett = $false
  foreach ($v in $vfHrar) {
    if ($SAMEINA_KODAR -contains $v.code) { if (-not $sett) { [void]$nyr.Add($m); $sett = $true } }
    else { [void]$nyr.Add($v) }
  }
  $vfHrar = $nyr
}

$vf_list = New-Object System.Collections.ArrayList
foreach ($l in $vfHrar) {
  $o = [ordered]@{ kg = $l.kg; code = $l.code; name = $l.name }
  $t = Til-Hlut $l
  foreach ($k in $t.Keys) { $o[$k] = $t[$k] }
  $st = New-Object System.Collections.ArrayList
  if ($stHrar.ContainsKey($l.code)) {
    foreach ($s in $stHrar[$l.code]) {
      $so = [ordered]@{ name = $s.name }
      $ts = Til-Hlut $s
      foreach ($k in $ts.Keys) { $so[$k] = $ts[$k] }
      [void]$st.Add($so)
    }
  }
  $o['stett'] = $st
  [void]$vf_list.Add($o)
}

# Samræmisprófun: summa viðfanga ≈ KG
foreach ($kode in $kg_totals.Keys) {
  $s = 0.0; foreach ($v in $vfHrar) { if ($v.kg -eq $kode) { $s += $v.medalt } }
  if ([Math]::Abs($s - $kgHrar[$kode].medalt) -gt 0.05) { Vidvorun ("{0}: summa viðfanga ({1:N2}) passar ekki við KG-samtölu ({2:N2})." -f $kode, $s, $kgHrar[$kode].medalt) }
}

# ── Yfirvinna ─────────────────────────────────────────────────────────────────
# Les blokkir: heiti (KG/viðfang) → raðir „Greidd dagvinnustöðugildi“, „Unnin og greidd yfirvinnustöðugildi“, „Yfirvinnuhlutfall“
function Lesa-YV($a, [string]$heitiMynstur) {
  $ut = [ordered]@{}
  $lc = Fyrsti-Textadalkur $a
  $nu = $null; $hd = $null
  for ($r = 1; $r -le (Rows $a); $r++) {
    $lab = Txt (Cell $a $r $lc)
    if ($lab -eq '') { continue }
    if ($heitiMynstur -and $lab -match $heitiMynstur) {
      $nu = $Matches[1]
      if (-not $ut.Contains($nu)) { $ut[$nu] = @{} } else { Vidvorun "Tvítekin yfirvinnublokk fyrir $nu — síðari hunsuð."; $nu = '__tvitekid__'; $ut[$nu] = @{} }
      continue
    }
    if ($lab -eq 'Stöðugildi') {
      $hd = @{ man = @{}; medalt = $null }
      for ($c = $lc + 1; $c -le (Cols $a); $c++) {
        $t = Txt (Cell $a $r $c)
        if ($t -match '^(0[1-9]|1[0-2])$') { $hd.man[$t] = $c } elseif ($t -match '^Meðaltal \d{4}$' -and $null -eq $hd.medalt) { $hd.medalt = $c }
      }
      continue
    }
    if (-not $hd) { continue }
    $f = switch ($lab) { 'Greidd dagvinnustöðugildi' { 'dv' } 'Unnin og greidd yfirvinnustöðugildi' { 'ot' } 'Yfirvinnuhlutfall' { 'yv' } default { $null } }
    if (-not $f) { continue }
    $k = $(if ($heitiMynstur) { $nu } else { '__svid__' })
    if (-not $k) { continue }
    if (-not $ut.Contains($k)) { $ut[$k] = @{} }
    if ($ut[$k].ContainsKey($f)) { continue }
    $o = @{ man = @{}; medalt = (N0 (Cell $a $r $hd.medalt)) }
    foreach ($mk in $manKodar) { $o.man[$mk] = $(if ($hd.man.ContainsKey($mk)) { N0 (Cell $a $r $hd.man[$mk]) } else { 0.0 }) }
    $ut[$k][$f] = $o
  }
  if ($ut.Contains('__tvitekid__')) { $ut.Remove('__tvitekid__') }
  return $ut
}
function YV-Hlutur($b) {
  $o = [ordered]@{}
  for ($i = 0; $i -lt $manKodar.Count; $i++) { $o[$manLyklar[$i]] = $(if ($b -and $b.yv) { Rnd $b.yv.man[$manKodar[$i]] 4 } else { 0.0 }) }
  $o['medalt'] = $(if ($b -and $b.yv) { Rnd $b.yv.medalt 4 } else { 0.0 })
  return $o
}

$yvSvid = Lesa-YV $YV['Yfirvinnugreining -Svið'] $null
$yvKG = Lesa-YV $YV['Yfirvinnugreining - sviðsein.'] '^(KG\d\d) '
$yvVF = Lesa-YV $YV['Yfirvinnugreining - viðföng'] '^(\d{5}) '

$kg_yv = [ordered]@{}
foreach ($k in $yvKG.Keys) { $kg_yv[$k] = YV-Hlutur $yvKG[$k] }
foreach ($k in $kg_totals.Keys) { if (-not $kg_yv.Contains($k)) { Vidvorun "Engin yfirvinnugögn fyrir $k." } }

$vf_yv = [ordered]@{}
$samSett = $false
foreach ($k in $yvVF.Keys) {
  if ($SAMEINA_KODAR -contains $k) {
    if ($samSett) { continue }
    $samSett = $true
    $bl = @($SAMEINA_KODAR | Where-Object { $yvVF.Contains($_) } | ForEach-Object { $yvVF[$_] })
    $o = [ordered]@{}
    for ($i = 0; $i -lt $manKodar.Count; $i++) {
      $mk = $manKodar[$i]; $ot = 0.0; $dv = 0.0
      foreach ($b in $bl) { if ($b.ot) { $ot += $b.ot.man[$mk] }; if ($b.dv) { $dv += $b.dv.man[$mk] } }
      $o[$manLyklar[$i]] = $(if ($dv -ne 0) { Rnd ($ot / $dv) 4 } else { 0.0 })
    }
    $ot = 0.0; $dv = 0.0
    foreach ($b in $bl) { if ($b.ot) { $ot += $b.ot.medalt }; if ($b.dv) { $dv += $b.dv.medalt } }
    $o['medalt'] = $(if ($dv -ne 0) { Rnd ($ot / $dv) 4 } else { 0.0 })
    $vf_yv[$SAMEINA_KODI] = $o
    continue
  }
  $vf_yv[$k] = YV-Hlutur $yvVF[$k]
}
$avg_yv = $(if ($yvSvid.Contains('__svid__') -and $yvSvid['__svid__'].yv) { Rnd $yvSvid['__svid__'].yv.medalt 4 } else { Vidvorun 'Yfirvinnuhlutfall sviðs fannst ekki.'; 0.0 })

# ── Veikindi ──────────────────────────────────────────────────────────────────
function VK-Dalkar($a, [int]$r) {
  $h = @{}
  for ($c = 1; $c -le (Cols $a); $c++) {
    $t = Txt (Cell $a $r $c)
    switch -Regex ($t) {
      '^Greidd stg' { $h.greidd = $c }
      '^Veikinda stg' { $h.veik = $c }
      '^Heildar hlutfall' { $h.heildar = $c }
      '^Skammtíma' { $h.skammt_stg = $c; $h.skammt = $c + 1 }
      '^Meðallöng' { $h.medal_stg = $c; $h.medal = $c + 1 }
      '^Langtíma' { $h.lang_stg = $c; $h.lang = $c + 1 }
    }
  }
  return $h
}
$A = $VK['Tímabilið']
$kg_vk = $null
for ($r = 1; $r -le (Rows $A); $r++) {
  $h = VK-Dalkar $A $r
  if ($h.greidd) {
    $kg_vk = [ordered]@{ greidd = Rnd (N0 (Cell $A ($r + 1) $h.greidd)) 2; heildar_pct = Rnd (N0 (Cell $A ($r + 1) $h.heildar)) 4; skammt_pct = Rnd (N0 (Cell $A ($r + 1) $h.skammt)) 4; lang_pct = Rnd (N0 (Cell $A ($r + 1) $h.lang)) 4 }
    break
  }
}
if (-not $kg_vk) { throw 'Fann ekki „Greidd stg“ í blaðinu „Tímabilið“.' }

$A = $VK['Skipulagseiningar']
$deild_vk = New-Object System.Collections.ArrayList
$h = $null; $samVK = @()
for ($r = 1; $r -le (Rows $A); $r++) {
  if (-not $h) { $t = VK-Dalkar $A $r; if ($t.greidd) { $h = $t }; continue }
  $nafn = Txt (Cell $A $r ($h.greidd - 1))
  if ($nafn -eq '') { continue }
  $o = [ordered]@{ name = $nafn; greidd = (N0 (Cell $A $r $h.greidd)); heildar_pct = (N0 (Cell $A $r $h.heildar)); skammt_pct = (N0 (Cell $A $r $h.skammt)); lang_pct = (N0 (Cell $A $r $h.lang)) }
  $o['_veik'] = N0 (Cell $A $r $h.veik); $o['_sk'] = N0 (Cell $A $r $h.skammt_stg); $o['_la'] = N0 (Cell $A $r $h.lang_stg)
  $erSam = $false; foreach ($p in $SAMEINA_VK) { if ($nafn -match $p) { $erSam = $true } }
  if ($erSam) { $samVK += , $o; if ($samVK.Count -eq 1) { [void]$deild_vk.Add('__SAM__') }; continue }
  # Sama sía og í frumriti: deildir með fleiri en 1 greitt stöðugildi (meðaltal) + meðaltalslína
  if ($o.greidd -gt 1 -or $nafn -eq 'Meðaltal:') { [void]$deild_vk.Add($o) }
  if ($nafn -eq 'Meðaltal:') { break }
}
$deildOut = New-Object System.Collections.ArrayList
foreach ($d in $deild_vk) {
  if ($d -is [string]) {
    $g = 0.0; $v = 0.0; $s = 0.0; $l = 0.0
    foreach ($x in $samVK) { $g += $x.greidd; $v += $x._veik; $s += $x._sk; $l += $x._la }
    if ($g -gt 1) { [void]$deildOut.Add([ordered]@{ name = $SAMEINA_VK_NAFN; greidd = Rnd $g 2; heildar_pct = Rnd ($v / $g) 4; skammt_pct = Rnd ($s / $g) 4; lang_pct = Rnd ($l / $g) 4 }) }
  } else {
    [void]$deildOut.Add([ordered]@{ name = $d.name; greidd = Rnd $d.greidd 2; heildar_pct = Rnd $d.heildar_pct 4; skammt_pct = Rnd $d.skammt_pct 4; lang_pct = Rnd $d.lang_pct 4 })
  }
}

# ── Samantekt (summa slétttaðra KG-gilda, eins og í frumriti) ───────────────────
$lk = $manLyklar[-1]
$sum = @{ last = 0.0; medalt = 0.0; med25 = 0.0; aatl = 0.0; frav = 0.0 }
foreach ($k in $kg_totals.Values) { $sum.last += $k[$lk]; $sum.medalt += $k.medalt26; $sum.med25 += $k.medalt25; $sum.aatl += $k.aatl; $sum.frav += $k.frav }
$_summary = [ordered]@{
  total_last = Rnd $sum.last 2
  total_medalt = Rnd $sum.medalt 2
  total_med25 = Rnd $sum.med25 2
  total_aatl = Rnd $sum.aatl 2
  total_frav = Rnd $sum.frav 2
  total_frav_pct = $(if ($sum.aatl -ne 0) { Rnd ($sum.frav / $sum.aatl) 4 } else { 0.0 })
  avg_yv = $avg_yv
  kg_colors = $KG_LITIR
}
$_summary.Insert(1, "total_$lk", (Rnd $sum.last 2))

# ── Meta ──────────────────────────────────────────────────────────────────────
$sidHeiti = $MAN_FULLT[$sidastiIdx]
$meta = [ordered]@{
  key = $KEY
  manudur = $Manudur
  ar = $AR
  fjoldi_manada = $manKodar.Count
  sidasti = $lk
  sidasti_heiti = $sidHeiti
  sidasti_heiti_stor = ($sidHeiti.Substring(0, 1).ToUpper() + $sidHeiti.Substring(1))
  timabil = ('{0}–{1} {2}' -f $MAN_STUTT[$fyrstiIdx].ToLower(), $MAN_STUTT[$sidastiIdx].ToLower(), $AR)
  titill = ('{0} {1}' -f ($sidHeiti.Substring(0, 1).ToUpper() + $sidHeiti.Substring(1)), $AR)
  heimildir = @($dvSkra.Name, $yvSkra.Name, $vkSkra.Name)
  buid_til = (Get-Date).ToString('yyyy-MM-dd HH:mm')
  vidvaranir = @($Vidvaranir)
}

$D = [ordered]@{
  meta = $meta
  months = $manLyklar
  month_labels = $manStutt
  kg_totals = $kg_totals
  vf_list = $vf_list
  kg_yv = $kg_yv
  vf_yv = $vf_yv
  kg_vk = $kg_vk
  deild_vk = $deildOut
  _summary = $_summary
}

# ── JSON-ritari (stýrt snið, óháð ConvertTo-Json í PS 5.1) ─────────────────────
function Js-Str([string]$s) {
  $sb = New-Object Text.StringBuilder
  [void]$sb.Append('"')
  foreach ($ch in $s.ToCharArray()) {
    switch ($ch) {
      '"' { [void]$sb.Append('\"') } '\' { [void]$sb.Append('\\') }
      "`n" { [void]$sb.Append('\n') } "`r" { [void]$sb.Append('\r') } "`t" { [void]$sb.Append('\t') }
      '<' { [void]$sb.Append('\u003c') }
      default { if ([int]$ch -lt 32) { [void]$sb.Append(('\u{0:x4}' -f [int]$ch)) } else { [void]$sb.Append($ch) } }
    }
  }
  [void]$sb.Append('"')
  return $sb.ToString()
}
function Js($v, [int]$dypt = 0) {
  if ($null -eq $v) { return 'null' }
  if ($v -is [string]) { return Js-Str $v }
  if ($v -is [bool]) { return $(if ($v) { 'true' } else { 'false' }) }
  if ($v -is [double] -or $v -is [single] -or $v -is [decimal]) {
    $d = [double]$v
    if ([double]::IsNaN($d) -or [double]::IsInfinity($d)) { return 'null' }
    return $d.ToString('R', $Inv)
  }
  if ($v -is [int] -or $v -is [long]) { return $v.ToString($Inv) }
  $inn = '  ' * ($dypt + 1); $ut = '  ' * $dypt
  if ($v -is [Collections.IDictionary]) {
    $p = @(); foreach ($k in $v.Keys) { $p += ($inn + (Js-Str "$k") + ': ' + (Js $v[$k] ($dypt + 1))) }
    if ($p.Count -eq 0) { return '{}' }
    # Smáhlutir (aðeins tölur/strengir) á einni línu
    $einfalt = $true; foreach ($k in $v.Keys) { $x = $v[$k]; if ($x -is [Collections.IDictionary] -or ($x -is [Collections.IEnumerable] -and $x -isnot [string])) { $einfalt = $false } }
    if ($einfalt -and $dypt -ge 1) { return '{' + (($v.Keys | ForEach-Object { (Js-Str "$_") + ': ' + (Js $v[$_] 0) }) -join ', ') + '}' }
    return "{`n" + ($p -join ",`n") + "`n$ut}"
  }
  if ($v -is [Collections.IEnumerable]) {
    $p = @(); foreach ($x in $v) { $p += (Js $x ($dypt + 1)) }
    if ($p.Count -eq 0) { return '[]' }
    $einfalt = $true; foreach ($x in $v) { if ($x -is [Collections.IDictionary] -or ($x -is [Collections.IEnumerable] -and $x -isnot [string])) { $einfalt = $false } }
    if ($einfalt) { return '[' + ($p -join ', ') + ']' }
    return "[`n" + (($p | ForEach-Object { $inn + $_ }) -join ",`n") + "`n$ut]"
  }
  return Js-Str "$v"
}

$haus = "// Gögn fyrir mælaborðið «Stöðugildi» — $($meta.timabil). Búið til sjálfvirkt af scripts/stodugildi.ps1 $($meta.buid_til).`r`n// Heimildir: $($meta.heimildir -join ' | ')`r`n"
$texti = $haus + 'window.D = ' + (Js $D 0) + ";`r`n"
$utMappa = Join-Path (Join-Path $Rot 'data') $Manudur
New-Item -ItemType Directory -Force -Path $utMappa | Out-Null
$utSkra = Join-Path $utMappa "$KEY.js"
[IO.File]::WriteAllText($utSkra, $texti, (New-Object Text.UTF8Encoding $false))
Write-Host "Skrifað: $utSkra"
Write-Host ("Mánuðir: {0} ({1}) · KG: {2} · viðföng: {3} · DV stg. {4}: {5} · meðaltal: {6} · áætlun: {7} · frávik: {8}" -f $manKodar.Count, $meta.timabil, $kg_totals.Count, $vf_list.Count, $sidHeiti, $_summary.total_last, $_summary.total_medalt, $_summary.total_aatl, $_summary.total_frav)
if ($Vidvaranir.Count) { Write-Host "Viðvaranir: $($Vidvaranir.Count)" }
