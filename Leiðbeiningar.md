# Mælaborð BLEND: mánaðarleg uppfærsla

## Opna appið
- **Á netinu (sími eða tölva):** https://arnija-lsh.github.io/maelabord-blend/
- **Á tölvunni:** tvísmelltu á `index.html`. Appið opnast í Edge eða Chrome.

Veldu tímabil efst til hægri og mælaborð í flipunum. Slóðin breytist eftir vali (t.d. `#yfirvinna/2026-09`), svo hægt er að senda tengil beint á tiltekið mælaborð.

## Að nota mælaborðin
- **Fjárhagur:** Töflurnar sýna tekjur (4), launagjöld (5A) skipt í föst (50) og breytileg laun (51), og önnur rekstrargjöld (5B). Smelltu á sviðseiningu til að sjá deildir. „📊 Greining“ sýnir brúarmynd og frávik kostnaðarliða. Með **+** við deild, eða við **BLEND Samtals** í Viðföng-flipanum, opnast brúarmynd frá áætlun til raunar. Litareglur og orðalisti eru í flipanum *Hjálp & orðalisti*.
- **Yfirvinna:** Kortin sýna yfirvinnuvaktir umfram viðmið, þ.e. 3,5% af mönnunarheimild. Yfirvinnustöðugildi (YV klst. ÷ 173,4) eru sýnd í kössum, kortum og töflum, og YTD sem meðaltal á mánuð. Í Viðföng-flipanum fylgja athugasemdirnar valinni sviðseiningu. Skilgreiningar eru í flipanum *Leiðbeiningar*.
- **Stöðugildi:** Efsti ramminn fylgir sviðseiningunni sem er valin í Viðföng-flipanum. Skýringar eru í flipanum *Um skýrsluna*.
- **Lyfjakostnaður:** Í Yfirliti er hægt að smella á kostnaðarlið, ATC-flokk, sviðseiningu eða mánuð. Hin gröfin sýna þá aðeins það sem er valið. **✕ Hreinsa val** fer aftur í heildina.

## Uppfæra með nýjum mánuði
1. Settu Excel-skrár mánaðarins í nýja möppu undir `Gögn`, eins og áður (t.d. `Gögn\Nóvember`).
   Skriftan þarf þessar skrár:
   - Heildaruppgjör
   - 5.1 Yfirvinnugreining
   - Yfirlit dagvinnustöðugilda (KG)
   - Veikindagreining fyrir svið
   - 8.1 og 8.2 Lyfjakaup
2. Opnaðu PowerShell í möppunni `Mælaborð BLEND` og keyrðu skipunina hér fyrir neðan.
   `-Manudur` er síðasti uppgjörsmánuðurinn, ekki heiti möppunnar.

   ```
   powershell -ExecutionPolicy Bypass -File .\Uppfaera.ps1 -Mappa "..\Gögn\Nóvember" -Manudur 2026-10
   ```

   Uppfærslan tekur um eina mínútu. Á meðan opnast Excel í bakgrunni.
3. **Uppgjörsskýringar (valfrjálst):** Afritaðu `skyringar.js` úr síðasta mánuði inn í nýju mánaðarmöppuna, t.d. `data\2026-10\`. Opnaðu hana í Notepad og uppfærðu línurnar með tölum úr „Uppgjör (4 af 5) – Helstu frávik YTD“. Fjárhæðir eru í krónum, með sama formerki og í kerfinu. Ef skráin er ekki til fyrir mánuðinn birtist skýringabrúin einfaldlega ekki.
4. Opnaðu `index.html`. Nýi mánuðurinn birtist í tímabilsvalinu og eldri mánuðir eru áfram aðgengilegir.
5. **Birta á netinu:** Skriftan opnar sjálf `data`-möppuna og upphleðslusíðu GitHub.
   - Dragðu nýju mánaðarmöppuna (t.d. `2026-10`) og skrána `manudir.js` inn á síðuna.
   - Smelltu á **Commit changes**.
   - Eftir 1–2 mínútur er nýi mánuðurinn kominn á https://arnija-lsh.github.io/maelabord-blend/
   - Ef þú breyttir útliti mælaborðs þarftu líka að hlaða upp viðkomandi skrá úr `dashboards`.
   - Ef GitHub sýnir „HTTP ERROR 400“ eftir **Commit changes** hefur upphleðslan samt yfirleitt tekist. Athugaðu að skráin sýni „now“ eða „1 minute ago“ í kóðageymslunni.
   - Til að eyða skrá á GitHub: opnaðu skrána, veldu **„…“ → Delete file** og staðfestu svo með græna hnappnum **Commit changes…** efst til hægri.

## Hvað er hvar
| Mappa eða skrá | Innihald |
|---|---|
| `index.html` | Appið sjálft (flipar og tímabilsval) |
| `dashboards\` | Útlit mælaborðanna fjögurra |
| `data\<mánuður>\` | Tölur hvers mánaðar, búnar til af skriftunum |
| `scripts\` | Skriftur sem lesa Excel-skrárnar, ein fyrir hvert mælaborð |
| `Uppfaera.ps1` | Keyrir allar skriftur og bætir mánuðinum í tímabilsvalið |

## Athugið
- Skýringartextar í mælaborðunum eru myndaðir sjálfkrafa út frá tölunum. Í fjárhagsmælaborðinu birtist gul athugasemd þegar fastir textar um gagnagæði þurfa yfirferð.
- Ef skráaheiti breytast mikið getur skrifta hætt að finna skrá. Þá birtist villa fyrir það mælaborð, en hin uppfærast samt.
- Appið sækir alltaf nýjustu útgáfu þegar skipt er um flipa. Ef gömul útgáfa birtist samt, ýttu einu sinni á **F5**.
