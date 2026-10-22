# Uppfærir öll mælaborð með gögnum nýs mánaðar.
#
# Dæmi (gögn fyrir jan–október 2026 liggja í Gögn\Nóvember):
#   .\Uppfaera.ps1 -Mappa "..\Gögn\Nóvember" -Manudur 2026-10
#
# -Manudur er síðasti uppgjörsmánuðurinn (ár-mánuður), ekki mánuðurinn sem mappan heitir eftir.
param(
  [Parameter(Mandatory = $true)][string]$Mappa,
  [Parameter(Mandatory = $true)][ValidatePattern('^\d{4}-\d{2}$')][string]$Manudur
)

$ErrorActionPreference = 'Stop'
$rot = $PSScriptRoot
$Mappa = (Resolve-Path $Mappa).Path

$manudir = 'janúar','febrúar','mars','apríl','maí','júní','júlí','ágúst','september','október','nóvember','desember'
$ar = [int]$Manudur.Substring(0, 4)
$nr = [int]$Manudur.Substring(5, 2)
$heiti = if ($nr -eq 1) { "Janúar $ar" } else { "Janúar–$($manudir[$nr - 1]) $ar" }

Write-Host "Uppfæri mælaborð fyrir $heiti úr $Mappa" -ForegroundColor Cyan

$villur = @()
foreach ($key in 'fjarhagur', 'yfirvinna', 'stodugildi', 'lyf') {
  Write-Host "  - $key ..." -NoNewline
  try {
    & (Join-Path $rot "scripts\$key.ps1") -Mappa $Mappa -Manudur $Manudur | Out-Null
    Write-Host " í lagi" -ForegroundColor Green
  } catch {
    Write-Host " VILLA: $($_.Exception.Message)" -ForegroundColor Red
    $villur += $key
  }
}

# Bæta mánuðinum í data\manudir.js ef hann er ekki þar fyrir.
$skra = Join-Path $rot 'data\manudir.js'
$texti = [IO.File]::ReadAllText($skra, [Text.Encoding]::UTF8)
if ($texti -notmatch "id: '$Manudur'") {
  $texti = $texti -replace "\r?\n\];", ",`r`n  { id: '$Manudur', heiti: '$heiti' }`r`n];"
  [IO.File]::WriteAllText($skra, $texti, (New-Object Text.UTF8Encoding $false))
  Write-Host "Bætti $heiti við mánaðalistann." -ForegroundColor Cyan
}

# Opna data-möppuna og upphleðslusíðu GitHub svo hægt sé að draga nýju gögnin beint inn.
$github = 'https://github.com/arnija-LSH/maelabord-blend/upload/main/data'
Write-Host ""
Write-Host "Birting á netinu: dragðu möppuna '$Manudur' og skrána 'manudir.js' úr data-möppunni" -ForegroundColor Cyan
Write-Host "inn á GitHub-síðuna sem opnast, og smelltu á 'Commit changes'." -ForegroundColor Cyan
Write-Host "Síðan uppfærist eftir 1-2 mínútur: https://arnija-lsh.github.io/maelabord-blend/" -ForegroundColor Cyan
Invoke-Item (Join-Path $rot 'data')
Start-Process $github

if ($villur.Count) {
  Write-Host "Mistókst fyrir: $($villur -join ', '). Hin mælaborðin eru uppfærð." -ForegroundColor Yellow
  exit 1
}
Write-Host "Lokið. Opnaðu index.html til að skoða." -ForegroundColor Green
