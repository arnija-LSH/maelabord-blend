# Uppfærir gögn yfirvinnumælaborðsins (dashboards/yfirvinna.html).
#
# Dæmi:
#   .\scripts\yfirvinna.ps1 -Mappa "..\Gögn\Október" -Manudur 2026-09
#
# Les:
#   * "5.1 Yfirvinnugreining ... KG.xlsx"  (blöð: Yfirvinnugreining -Svið / - sviðsein. / - viðföng)
#   * "Yfirlit dagvinnustöðugilda ... KG.xlsx" (blöð: Dv.stg - mán / Dv.stg. - sviðsein. / Dv.stg. - viðföng)
# Skrifar: data\<Manudur>\yfirvinna.js  (window.D = {...};)
param(
  [Parameter(Mandatory = $true)][string]$Mappa,
  [Parameter(Mandatory = $true)][string]$Manudur
)

$ErrorActionPreference = 'Stop'
$rot   = Split-Path -Parent $PSScriptRoot
$Mappa = (Resolve-Path -LiteralPath $Mappa).Path
$IS    = [Globalization.CultureInfo]::GetCultureInfo('is-IS')
$INV   = [Globalization.CultureInfo]::InvariantCulture

# ── Fastar ─────────────────────────────────────────────────────────────────────
$VID = 8.0; $HLUT = 0.035; $VMIN = 173.4
$MN_ALL     = 'Jan','Feb','Mar','Apr','Maí','Jún','Júlí','Ágú','Sept','Okt','Nóv','Des'
$MNFULL_ALL = 'Janúar','Febrúar','Mars','Apríl','Maí','Júní','Júlí','Ágúst','September','Október','Nóvember','Desember'
$KGINFO = [ordered]@{
  'KG01' = @{ name = 'Skrifstofa BLEND';                           color = '#6b7280' }
  'KG02' = @{ name = 'Sérgreinar bráðaþjónustu og lyflækninga';    color = '#7c3aed' }
  'KG03' = @{ name = 'Ferlideildir';                               color = '#2563eb' }
  'KG04' = @{ name = 'Legudeildir';                                color = '#8B0000' }
  'KG05' = @{ name = 'Endurhæfing';                                color = '#059669' }
  'KG06' = @{ name = 'Þjálfunareiningar endurhæfingar';            color = '#d97706' }
  'KG07' = @{ name = 'Sérgreinar lyflækninga og endurhæfingar';    color = '#C0392B' }
}
$SAMEINA = @{ '11135' = '11131' }   # 11135 fer inn í 11131
$SAMEINA_NAFN = '11131+11135 Legudeild Minnisraskana / Líknardeild'

# ── Hjálparföll ───────────────────────────────────────────────────────────────
# Námundun eins og í upprunalegu greiningunni (Python round): helmingur að sléttri tölu,
# miðað við raunverulegt tvíundagildi (t.d. 75,35 -> 75,3 því það er geymt sem 75,3499...).
function Rnd([double]$x, [int]$d = 0) {
  $m = [decimal]::Parse($x.ToString('G17', [Globalization.CultureInfo]::InvariantCulture), [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture)
  [double][Math]::Round($m, $d, [MidpointRounding]::ToEven)
}
function Num($x) {
  if ($null -eq $x) { return $null }
  if ($x -is [double] -or $x -is [int] -or $x -is [decimal]) { return [double]$x }
  $s = "$x".Trim(); if ($s -eq '') { return $null }
  $d = 0.0
  if ([double]::TryParse($s, [Globalization.NumberStyles]::Float, $INV, [ref]$d)) { return $d }
  return $null
}
function Txt($x) { if ($null -eq $x) { '' } else { "$x".Trim() } }
function Vm([double]$dv) { $dv * $VMIN * $HLUT / $VID }
function F1($x) { ([double]$x).ToString('N1', $IS) }
function F0($x) { ([double]$x).ToString('N0', $IS) }

function Finna-Skra([string[]]$mynstur, [string]$lysing) {
  $c = @(Get-ChildItem -LiteralPath $Mappa -File | Where-Object {
      $n = $_.Name; ($n -notlike '~$*') -and ($n -like '*.xls*') -and (@($mynstur | Where-Object { $n -like $_ }).Count -gt 0) })
  if (-not $c.Count) { throw "Fann ekki $lysing í $Mappa (mynstur: $($mynstur -join ', '))" }
  # KG-skrá tekin fram yfir aðrar, svo nýjasta
  return ($c | Sort-Object @{ e = { if ($_.Name -match 'KG') { 0 } else { 1 } } }, @{ e = 'LastWriteTime'; Descending = $true } | Select-Object -First 1)
}

function JsStr([string]$s) {
  $sb = New-Object Text.StringBuilder; [void]$sb.Append('"')
  foreach ($ch in $s.ToCharArray()) {
    switch ($ch) {
      '"'  { [void]$sb.Append('\"') } '\' { [void]$sb.Append('\\') }
      "`n" { [void]$sb.Append('\n') } "`r" { [void]$sb.Append('\r') } "`t" { [void]$sb.Append('\t') }
      default { if ([int]$ch -lt 32) { [void]$sb.Append(('\u{0:x4}' -f [int]$ch)) } else { [void]$sb.Append($ch) } }
    }
  }
  [void]$sb.Append('"'); $sb.ToString()
}
function ToJs($o) {
  if ($null -eq $o) { return 'null' }
  if ($o -is [string]) { return (JsStr $o) }
  if ($o -is [bool]) { if ($o) { return 'true' } else { return 'false' } }
  if ($o -is [double] -or $o -is [int] -or $o -is [long] -or $o -is [decimal]) {
    $d = [double]$o; if ($d -eq 0) { $d = 0.0 }   # engin -0
    return $d.ToString('R', $INV)
  }
  if ($o -is [Collections.IDictionary]) {
    $p = foreach ($k in $o.Keys) { (JsStr "$k") + ': ' + (ToJs $o[$k]) }
    return '{' + ($p -join ', ') + '}'
  }
  if ($o -is [Collections.IEnumerable]) {
    $p = foreach ($x in $o) { ToJs $x }
    return '[' + ($p -join ', ') + ']'
  }
  return (JsStr "$o")
}

# Les blokkir úr 5.1-blaði: hver blokk byrjar á línu með "KGxx ..." eða "nnnnn ..." í fyrsta dálki.
function Lesa-Yfirvinnublad($v, [string]$hausMynstur) {
  $nr = $v.GetUpperBound(0); $nc = $v.GetUpperBound(1)
  $blokkir = New-Object Collections.ArrayList
  $kg = ''; $cur = $null; $manCol = @{}; $medCol = 0
  for ($r = 1; $r -le $nr; $r++) {
    $a = Txt $v[$r, 1]
    if ($a -eq '') { continue }
    if ($a -match '^(KG\d\d)\s') { $kg = $Matches[1] }
    if ($a -match $hausMynstur) {
      $cur = [ordered]@{ kg = $kg; heiti = $a; yv = @{}; dv = @{}; medal = $null }
      [void]$blokkir.Add($cur); continue
    }
    if ($null -eq $cur) { continue }
    if ($a -eq 'Yfirvinnugreining' -or $a -eq 'Stöðugildi') {
      $manCol = @{}; $medCol = 0
      for ($c = 2; $c -le $nc; $c++) {
        $h = Txt $v[$r, $c]
        if ($h -match '^\d{1,2}$') { $manCol[[int]$h] = $c }
        elseif ($h -match '^(Meðaltal|Samtals)\s+\d{4}$' -and -not $medCol) { $medCol = $c }
      }
      continue
    }
    if ($a -like 'Samtals yfirvinnueiningar*') {
      foreach ($m in $manCol.Keys) { $x = Num $v[$r, $manCol[$m]]; if ($null -ne $x) { $cur.yv[$m] = $x } }
    } elseif ($a -like 'Greidd dagvinnust*') {
      foreach ($m in $manCol.Keys) { $x = Num $v[$r, $manCol[$m]]; if ($null -ne $x) { $cur.dv[$m] = $x } }
      if ($medCol) { $cur.medal = Num $v[$r, $medCol] }
    }
  }
  return , $blokkir
}

# Les "Áætlun 2026" úr stöðugildayfirliti: skilar hashtable lykill -> áætlun (heildarlínur, dálkur 2 auður).
function Lesa-Aetlun($v, [string]$lykilMynstur) {
  $nr = $v.GetUpperBound(0); $nc = $v.GetUpperBound(1)
  $ut = @{}; $aCol = 0
  for ($r = 1; $r -le $nr; $r++) {
    for ($c = 1; $c -le $nc; $c++) { if ((Txt $v[$r, $c]) -match '^Áætlun\s+\d{4}$') { $aCol = $c } }
    $a = Txt $v[$r, 1]; $b = Txt $v[$r, 2]
    if ($aCol -and $a -match $lykilMynstur -and $b -eq '') {
      $k = $Matches[1]; if (-not $ut.ContainsKey($k)) { $ut[$k] = Num $v[$r, $aCol] }
    }
  }
  return $ut
}
function Lesa-AetlunHeild($v) {
  $nr = $v.GetUpperBound(0); $nc = $v.GetUpperBound(1); $aCol = 0
  for ($r = 1; $r -le $nr; $r++) {
    for ($c = 1; $c -le $nc; $c++) {
      $t = Txt $v[$r, $c]
      if ($t -match '^Áætlun\s+\d{4}$') { $aCol = $c }
      if ($t -like 'Meðaltal*' -and $aCol) { return (Num $v[$r, $aCol]) }
    }
  }
  return $null
}

# ── Finna skrár ───────────────────────────────────────────────────────────────
$skraYv = Finna-Skra @('*Yfirvinnugreining*') 'yfirvinnugreiningu (5.1)'
$skraDv = Finna-Skra @('Yfirlit dagvinnust*') 'yfirlit dagvinnustöðugilda'
Write-Host "Yfirvinna: $($skraYv.Name)"
Write-Host "Stöðugildi: $($skraDv.Name)"

# Fyrra ár til samanburðar (valfrjálst): "5.1 Yfirvinnugreining ... <fyrra ár>.xlsx" í mánaðarmöppunni eða Gögn\Sameiginlegt.
# Sama uppbygging og 2026-skráin; þaðan koma bæði yfirvinnutímar og greidd DV stöðugildi fyrra árs.
$fyrraAr = $(if ($Manudur -match '^(\d{4})') { [int]$Matches[1] - 1 } else { (Get-Date).Year - 1 })
$skraFyrra = @(@($Mappa, (Join-Path (Split-Path -Parent $Mappa) 'Sameiginlegt')) | Where-Object { Test-Path -LiteralPath $_ } |
  ForEach-Object { Get-ChildItem -LiteralPath $_ -File } |
  Where-Object { $_.Name -notlike '~$*' -and $_.Name -like '*Yfirvinnugreining*' -and $_.Name -match "\b$fyrraAr\b" -and $_.Name -notmatch 'jan-' } |
  Sort-Object LastWriteTime -Descending) | Select-Object -First 1
if ($skraFyrra) { Write-Host "Fyrra ár ($fyrraAr): $($skraFyrra.Name)" } else { Write-Warning "Engin yfirvinnugreining fyrir $fyrraAr fannst; samanburður við fyrra ár sleppt." }

# ── Lesa Excel ────────────────────────────────────────────────────────────────
$blod = @{}
$xl = $null
try {
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
  $parList = @(@($skraYv.FullName, @{ svid = 'Yfirvinnugreining\s*-?\s*Svi[ðd]$'; svidein = 'Yfirvinnugreining.*svi[ðd]sein'; vidf = 'Yfirvinnugreining.*vi[ðd]f' }),
               @($skraDv.FullName, @{ dvman = '^Dv\.stg\s*-\s*mán$'; dvkg = '^Dv\.stg\.\s*-\s*svi[ðd]sein'; dvvidf = '^Dv\.stg\.\s*-\s*vi[ðd]f' }))
  if ($skraFyrra) { $parList += , @($skraFyrra.FullName, @{ f_svid = 'Yfirvinnugreining\s*-?\s*Svi[ðd]$'; f_svidein = 'Yfirvinnugreining.*svi[ðd]sein'; f_vidf = 'Yfirvinnugreining.*vi[ðd]f' }) }
  foreach ($par in $parList) {
    $wb = $null
    try {
      $wb = $xl.Workbooks.Open($par[0], 0, $true)
      foreach ($ws in $wb.Worksheets) {
        foreach ($k in $par[1].Keys) {
          if (-not $blod.ContainsKey($k) -and $ws.Name.Trim() -match $par[1][$k]) {
            $ur = $ws.UsedRange; $blod[$k] = $ur.Value2
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ur)
          }
        }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws)
      }
    } finally {
      if ($wb) { $wb.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
    }
  }
} finally {
  if ($xl) { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
  $xl = $null; [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
foreach ($k in 'svid', 'svidein', 'vidf', 'dvman', 'dvkg') { if (-not $blod.ContainsKey($k)) { throw "Blað fannst ekki: $k" } }

# ── Svið (BLEND heild) ────────────────────────────────────────────────────────
$svid = (Lesa-Yfirvinnublad $blod['svid'] '^KG\s')[0]
$nm = 0
foreach ($m in 1..12) { if ($svid.yv.ContainsKey($m) -or $svid.dv.ContainsKey($m)) { $nm = $m } }
if ($nm -lt 1) { throw 'Engir mánuðir með gögnum í yfirvinnugreiningu.' }
$L = $nm - 1
$man = 1..$nm
function Arr($h, [int]$d) { , @($man | ForEach-Object { $x = $h[$_]; if ($null -eq $x) { $x = 0 }; Rnd $x $d }) }

$aetlunKg = Lesa-Aetlun $blod['dvkg'] '^(KG\d\d)\s'
$aetlunHeild = Lesa-AetlunHeild $blod['dvman']
$aetlunVidf = @{}; if ($blod.ContainsKey('dvvidf')) { $aetlunVidf = Lesa-Aetlun $blod['dvvidf'] '^(\d{5})\s' }

$D = [ordered]@{}
$D.blend = [ordered]@{ yv = (Arr $svid.yv 0); dv = (Arr $svid.dv 2); aetlun = (Rnd $aetlunHeild 2) }

# ── Sviðseiningar ─────────────────────────────────────────────────────────────
$kgBlokkir = Lesa-Yfirvinnublad $blod['svidein'] '^KG\d\d\s'
$D.kg = [ordered]@{}
$kgRaw = @{}
foreach ($b in $kgBlokkir) {
  $k = $b.kg
  if ($D.kg.Contains($k)) { Write-Warning "Tvítekin sviðseining $k hunsuð"; continue }
  $info = $KGINFO[$k]; if (-not $info) { $info = @{ name = ($b.heiti -replace '^KG\d\d\s+', ''); color = '#2E75B6' } }
  $D.kg[$k] = [ordered]@{ name = $info.name; color = $info.color; yv = (Arr $b.yv 0); dv = (Arr $b.dv 2); aetlun = (Rnd ([double]$aetlunKg[$k]) 2) }
  $kgRaw[$k] = $b
}
# raða KG01..KG07
$tmp = [ordered]@{}; foreach ($k in ($D.kg.Keys | Sort-Object)) { $tmp[$k] = $D.kg[$k] }; $D.kg = $tmp

# ── Viðföng ───────────────────────────────────────────────────────────────────
$vBlokkir = Lesa-Yfirvinnublad $blod['vidf'] '^\d{5}\s'
$einingar = [ordered]@{}
foreach ($b in $vBlokkir) {
  $code = $b.heiti.Substring(0, 5)
  $lyk = $code; if ($SAMEINA.ContainsKey($code)) { $lyk = $SAMEINA[$code] }
  if (-not $einingar.Contains($lyk)) {
    $einingar[$lyk] = [ordered]@{ code = $lyk; kg = $b.kg; heiti = $b.heiti; yv = @{}; dv = @{}; medal = 0.0; hlutar = @{} }
  }
  $e = $einingar[$lyk]
  if ($e.hlutar.ContainsKey($code)) { Write-Warning "Tvítekið viðfang $code hunsað"; continue }   # Oracle-tvítekning
  $e.hlutar[$code] = 1
  foreach ($m in $b.yv.Keys) { $e.yv[$m] = [double]$e.yv[$m] + $b.yv[$m] }
  foreach ($m in $b.dv.Keys) { $e.dv[$m] = [double]$e.dv[$m] + $b.dv[$m] }
  if ($null -ne $b.medal) { $e.medal += $b.medal }
}
if ($einingar.Contains('11131') -and $einingar['11131'].hlutar.Count -gt 1) { $einingar['11131'].heiti = $SAMEINA_NAFN }

$vidfong = New-Object Collections.ArrayList
$extra = @{}   # mánaðarleg gögn per viðfang fyrir skýringartexta
foreach ($e in $einingar.Values) {
  $yvRaw = @($man | ForEach-Object { [double]$e.yv[$_] })
  $dvRaw = @($man | ForEach-Object { [double]$e.dv[$_] })
  $ytdRaw = ($yvRaw | Measure-Object -Sum).Sum
  # DV stöðugildi (mánaðar og meðaltal) eru námunduð í 2 aukastafi áður en reiknað er, eins og í D.kg
  $medal = Rnd $e.medal 2
  $dv7 = Rnd $dvRaw[$L] 2
  $aetl = Rnd $medal 1
  # engin starfsemi á árinu (engin yfirvinna og meðal-DV < 0,05) -> sleppt
  if ((Rnd $ytdRaw 0) -eq 0 -and (Rnd $e.medal 1) -eq 0 -and (Rnd $dv7 1) -eq 0) { continue }
  $rv7 = Rnd ($yvRaw[$L] / $VID) 1
  $vma = Rnd (Vm $medal) 1
  $vmr = Rnd (Vm $dv7) 1
  $o = [ordered]@{
    code = $e.code; name = $e.heiti; kg = $e.kg
    yv = @($yvRaw | ForEach-Object { Rnd $_ 0 })
    dv = @($dvRaw | ForEach-Object { Rnd $_ 2 })   # mánaðarleg DV stöðugildi, notuð í greiningu deildar
    dv_jul = (Rnd $dv7 1); aetlun = $aetl; rv_jul = $rv7; vma = $vma; vmr = $vmr
    dv_a = (Rnd ($rv7 - $vma) 1); dv_r = (Rnd ($rv7 - $vmr) 1)
    yv_pct = $(if ($dv7 -gt 0) { Rnd ($yvRaw[$L] / $VMIN / $dv7 * 100) 1 } else { 0.0 })
    ytd_yv = (Rnd $ytdRaw 0)
    ytd_dv = (Rnd ($ytdRaw / $VID - $nm * $vma) 1)
  }
  [void]$vidfong.Add($o)
  $extra[$e.code] = @{ yv = $yvRaw; dv = $dvRaw; aetlun = $aetlunVidf[$e.code]; hlutar = @($e.hlutar.Keys) }
}
# ── Fyrra ár (12 mánuðir) fyrir svið, sviðseiningar og viðföng ─────────────────
function Arr12($h, [int]$d) { , @(1..12 | ForEach-Object { $x = $h[$_]; if ($null -eq $x) { $x = 0 }; Rnd $x $d }) }
if ($skraFyrra -and $blod.ContainsKey('f_svid')) {
  $sF = (Lesa-Yfirvinnublad $blod['f_svid'] '^KG\s')[0]
  $D.blend.fyrra = [ordered]@{ yv = (Arr12 $sF.yv 0); dv = (Arr12 $sF.dv 2) }
  foreach ($b in (Lesa-Yfirvinnublad $blod['f_svidein'] '^KG\d\d\s')) {
    if ($D.kg.Contains($b.kg) -and -not $D.kg[$b.kg].Contains('fyrra')) { $D.kg[$b.kg].fyrra = [ordered]@{ yv = (Arr12 $b.yv 0); dv = (Arr12 $b.dv 2) } }
  }
  # Viðföng: sömu sameiningarreglur og fyrir árið í ár; viðföng sem ekki eru til í ár detta út
  $vF = @{}
  foreach ($b in (Lesa-Yfirvinnublad $blod['f_vidf'] '^\d{5}\s')) {
    $code = $b.heiti.Substring(0, 5); $lyk = $code; if ($SAMEINA.ContainsKey($code)) { $lyk = $SAMEINA[$code] }
    if (-not $vF.ContainsKey($lyk)) { $vF[$lyk] = @{ yv = @{}; dv = @{}; hlutar = @{} } }
    if ($vF[$lyk].hlutar.ContainsKey($code)) { continue }; $vF[$lyk].hlutar[$code] = 1
    foreach ($m in $b.yv.Keys) { $vF[$lyk].yv[$m] = [double]$vF[$lyk].yv[$m] + $b.yv[$m] }
    foreach ($m in $b.dv.Keys) { $vF[$lyk].dv[$m] = [double]$vF[$lyk].dv[$m] + $b.dv[$m] }
  }
  $ann = 0
  foreach ($o in $vidfong) {
    if ($vF.ContainsKey($o.code)) { $o['fyrra'] = [ordered]@{ yv = (Arr12 $vF[$o.code].yv 0); dv = (Arr12 $vF[$o.code].dv 2) }; $ann++ }
  }
  Write-Host ("Fyrra ár: BLEND YV jan–{0} {1} klst. · {2} af {3} viðföngum með samanburð" -f $MN_ALL[$L].ToLower(), (F0 (($D.blend.fyrra.yv[0..$L] | Measure-Object -Sum).Sum)), $ann, $vidfong.Count)
}

$D.vidfong = @($vidfong | Sort-Object @{ e = { $_.kg } }, @{ e = { $_.ytd_yv }; Descending = $true }, @{ e = { $_.code } })

# ── Lýsigögn og textar ────────────────────────────────────────────────────────
$ar = 0
$vs = $blod['svid']
for ($r = 1; $r -le $vs.GetUpperBound(0) -and -not $ar; $r++) {
  for ($c = 1; $c -le $vs.GetUpperBound(1); $c++) { if ((Txt $vs[$r, $c]) -match '^Samtals\s+(\d{4})$') { $ar = [int]$Matches[1]; break } }
}
if (-not $ar -and $Manudur -match '^(\d{4})') { $ar = [int]$Matches[1] }
if (-not $ar) { $ar = (Get-Date).Year }
$mnLow = $MNFULL_ALL[$L].ToLower()
$mnCap = $MNFULL_ALL[$L]
$stuttSidasti = $MN_ALL[$L].ToLower()
$MN = @($MN_ALL[0..$L]); $MNFULL = @($MNFULL_ALL[0..$L])
$timabilStutt = if ($nm -eq 1) { 'jan' } else { "jan–$stuttSidasti" }
$yp = { param($y, $d) if ($d -gt 0) { $y / $VMIN / $d * 100 } else { 0 } }

$bY = $D.blend.yv; $bD = $D.blend.dv
$ypM = @(0..$L | ForEach-Object { Rnd (& $yp $bY[$_] $bD[$_]) 1 })
$ytdPct = Rnd (& $yp (($bY | Measure-Object -Sum).Sum / $nm) (($bD | Measure-Object -Sum).Sum / $nm)) 1

# Þrjár sviðseiningar með hæsta YTD YV% fara í þróunargraf
$kgPct = foreach ($k in $D.kg.Keys) {
  $kd = $D.kg[$k]
  [pscustomobject]@{ k = $k; p = (& $yp (($kd.yv | Measure-Object -Sum).Sum / $nm) (($kd.dv | Measure-Object -Sum).Sum / $nm)) }
}
$hiKgs = @($kgPct | Sort-Object p -Descending | Select-Object -First 3 | ForEach-Object { $_.k } | Sort-Object)

# Athugasemd við þróunargraf (sumarmánuðir)
$throun = ''
if ($nm -ge 6) {
  $sum = @(5..([Math]::Min($L, 7)))
  $sumHeiti = if ($sum.Count -eq 1) { $MNFULL_ALL[5] } else { "$($MNFULL_ALL[5])–$($MNFULL_ALL[$sum[-1]].ToLower())" }
  $sumP = @($sum | ForEach-Object { $ypM[$_] })
  $vor = @(0..4 | ForEach-Object { $ypM[$_] }); $vorM = Rnd (($vor | Measure-Object -Average).Average) 1
  $minS = ($sumP | Measure-Object -Minimum).Minimum; $maxS = ($sumP | Measure-Object -Maximum).Maximum
  $bil = if ($minS -eq $maxS) { "$(F1 $minS)%" } else { "$(F1 $minS)–$(F1 $maxS)%" }
  # mest DV-aukning frá maí að hæsta sumarmánuði
  $auk = foreach ($k in $D.kg.Keys) { $kd = $D.kg[$k]; $mx = (@($sum | ForEach-Object { $kd.dv[$_] }) | Measure-Object -Maximum).Maximum; [pscustomobject]@{ k = $k; a = $mx - $kd.dv[4] } }
  $topAuk = @($auk | Sort-Object a -Descending | Select-Object -First 3 | Sort-Object k | ForEach-Object { $_.k })
  $throun = "&#9728; <strong>$($sumHeiti):</strong> Yfirvinnuhlutfall BLEND var $bil í sumarmánuðum en $(F1 $vorM)% að meðaltali jan–maí — vænt mynstur vegna sumarleyfa, ekki endilega framför. DV stöðugildi hækkuðu mest frá maí í $($topAuk -join ', ')."
  if ($L -gt 7) { $throun += " $($mnCap): $(F1 $ypM[$L])%." }
}

# Viðvaranir
$alerts = New-Object Collections.ArrayList
foreach ($v in $D.vidfong) {
  if ($v.yv_pct -gt 30 -or $v.code -eq '09083') {
    $x = $extra[$v.code]
    $yfir100 = @(0..$L | Where-Object { $x.dv[$_] -gt 0 -and (& $yp $x.yv[$_] $x.dv[$_]) -gt 100 } | ForEach-Object { $MNFULL_ALL[$_].ToLower() })
    $t = "<strong>$($v.name) ($($v.kg)):</strong> YV% $(F1 $v.yv_pct)% í $mnLow ($(F0 $v.yv[$L]) klst. á $(F1 $v.dv_jul) DV stg.)."
    if ($yfir100.Count) { $t += " YV% yfir 100% í: $($yfir100 -join ', ')." }
    $t += ' Skoða VS-skráningar.'
    [void]$alerts.Add([ordered]@{ cls = 'red'; txt = $t })
  }
}
foreach ($v in $D.vidfong) {
  $x = $extra[$v.code]
  $a = $x.aetlun
  if ($v.ytd_yv -gt 0 -and ($null -eq $a -or $a -eq 0) -and $x.hlutar.Count -eq 1) {
    [void]$alerts.Add([ordered]@{ cls = 'warn'; txt = "<strong>$($v.name) ($($v.kg)):</strong> Engin áætluð DV stöðugildi 2026 í stöðugildayfirliti en $(F0 $v.ytd_yv) YV klst. $timabilStutt." })
  }
}
$iMin = 0; $iMax = 0
for ($i = 0; $i -le $L; $i++) { if ($ypM[$i] -lt $ypM[$iMin]) { $iMin = $i }; if ($ypM[$i] -gt $ypM[$iMax]) { $iMax = $i } }
[void]$alerts.Add([ordered]@{ cls = 'grn'; txt = "<strong>BLEND $($timabilStutt):</strong> Yfirvinnuhlutfall lægst í $($MNFULL_ALL[$iMin].ToLower()) ($(F1 $ypM[$iMin])%) og hæst í $($MNFULL_ALL[$iMax].ToLower()) ($(F1 $ypM[$iMax])%). YTD $(F1 $ytdPct)% (viðmið 3,5%)." })
$D.alerts = @($alerts)

$D.meta = [ordered]@{
  manudur_id   = $Manudur
  nm           = $nm
  ar           = $ar
  MN           = $MN
  MNFULL       = $MNFULL
  titill       = $(if ($nm -eq 1) { "Jan $ar" } else { "Jan–$($MN_ALL[$L].ToLower()) $ar" })
  badge        = $(if ($nm -eq 1) { "Janúar $ar" } else { "Janúar–$($mnCap.ToLower()) $ar" })
  langt        = $(if ($nm -eq 1) { "Janúar $ar" } else { "Janúar til $($mnCap.ToLower()) $ar" })
  timabil      = "$timabilStutt $ar"
  stutt        = $timabilStutt
  manudur      = $mnLow
  manudur_stor = $mnCap
  ytd_label    = "YTD $($timabilStutt.ToUpper())."
  hiKgs        = $hiKgs
  throun_note  = $throun
  fn_svid      = "Viðmið 3,5% · 8 klst./vakt · 173,4 klst./mán. · $mnCap-gögn eru bráðabirgðatölur · Ópersónugreinanlegt — LSH Fjármálasvið"
  gogn_lokad   = "$mnCap $ar — tölur eru bráðabirgðatölur"
  heimildir    = @(@($skraYv.Name, $skraDv.Name) + @($(if ($skraFyrra) { $skraFyrra.Name })) | Where-Object { $_ })
  fyrraAr      = $(if ($D.blend.Contains('fyrra')) { $fyrraAr } else { $null })
  sjalfvirkt   = 'Textar í throun_note, alerts, fn_svid og gogn_lokad eru myndaðir sjálfvirkt út frá gögnum — yfirfara.'
}
# meta fyrst
$Dut = [ordered]@{ meta = $D.meta }; foreach ($k in 'blend', 'kg', 'vidfong', 'alerts') { $Dut[$k] = $D[$k] }

# ── Skrifa ────────────────────────────────────────────────────────────────────
$utMappa = Join-Path $rot "data\$Manudur"
New-Item -ItemType Directory -Force -Path $utMappa | Out-Null
$utSkra = Join-Path $utMappa 'yfirvinna.js'
$js = "// Myndað af scripts/yfirvinna.ps1 $(Get-Date -Format 'yyyy-MM-dd HH:mm') úr: $($skraYv.Name); $($skraDv.Name)`r`nwindow.D = " + (ToJs $Dut) + ";`r`n"
[IO.File]::WriteAllText($utSkra, $js, (New-Object Text.UTF8Encoding $false))
Write-Host "Skrifaði $utSkra ($nm mánuðir, $($D.vidfong.Count) viðföng)"
