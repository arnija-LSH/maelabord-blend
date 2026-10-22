# Mælaborð BLEND: mánaðarleg uppfærsla

## Opna appið
Tvísmelltu á `index.html`. Appið opnast í Edge eða Chrome. Veldu tímabil efst til hægri og mælaborð í flipunum.

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
3. Opnaðu `index.html`. Nýi mánuðurinn birtist í tímabilsvalinu og eldri mánuðir eru áfram aðgengilegir.
4. **Birta á netinu:** Skriftan opnar sjálf `data`-möppuna og upphleðslusíðu GitHub.
   - Dragðu nýju mánaðarmöppuna (t.d. `2026-10`) og skrána `manudir.js` inn á síðuna.
   - Smelltu á **Commit changes**.
   - Eftir 1–2 mínútur er nýi mánuðurinn kominn á https://arnija-lsh.github.io/maelabord-blend/
   - Ef þú breyttir útliti mælaborðs þarftu líka að hlaða upp viðkomandi skrá úr `dashboards`.

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
