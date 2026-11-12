# Maßeinheiten im Dokumentmodell

Das `DocumentModel` speichert Maße in den Einheiten, die WordprocessingML auf der Festplatte verwendet. Seit die Einheiten-Brands eingeführt wurden, sind diese Einheiten Teil der TypeScript-Typen. Ein Schriftgrad, der dort übergeben wird, wo eine Länge erwartet wird, oder ein nicht gerundeter Pixelwert, wo ganze Twips verlangt sind, ist deshalb ein Kompilierfehler und keine beschädigte Datei.

Brands gibt es nur zur Kompilierzeit. Ein `Twips` ist zur Laufzeit eine einfache JavaScript-Zahl und bleibt dem Typ `number` zuweisbar. Arithmetik, Vergleiche und JSON-Serialisierung bleiben daher unverändert.

## Brands {#brands}

Alle Brands werden aus `docx-core` exportiert.

| Typ            | Einheit                          | Schema-Typ              | Konstruktor (strikt) | Rundender Konstruktor  |
| -------------- | -------------------------------- | ----------------------- | -------------------- | ---------------------- |
| `Twips`        | 1/20 Punkt, nicht negativ        | `ST_TwipsMeasure`       | `twips(n)`           | `roundTwips(n)`        |
| `SignedTwips`  | 1/20 Punkt, auch negativ möglich | `ST_SignedTwipsMeasure` | `signedTwips(n)`     | `roundSignedTwips(n)`  |
| `HalfPoints`   | 1/2 Punkt (Schriftgröße `w:sz`)  | `ST_HpsMeasure`         | `halfPoints(n)`      | `roundHalfPoints(n)`   |
| `EighthPoints` | 1/8 Punkt (Rahmenbreite)         | `ST_EighthPointMeasure` | `eighthPoints(n)`    | `roundEighthPoints(n)` |
| `Emu`          | 1/914400 Zoll (DrawingML)        | `ST_Coordinate`         | `emu(n)`             | `roundEmu(n)`          |

`Twips` ist ein Untertyp von `SignedTwips`: Ein vorzeichenloses Maß kann einem vorzeichenbehafteten Feld zugewiesen werden, aber nicht umgekehrt.

**Strikte Konstruktoren** lösen einen `RangeError` aus, sofern der Wert nicht bereits eine ganze, sichere Ganzzahl ist (und bei `twips` nicht negativ). Parser und Tests verwenden sie.

**Rundende Konstruktoren** sind für berechnete Werte gedacht (Lineal-Ziehen, Zoomberechnungen, Pixelumrechnungen). Sie runden auf die nächste ganze Zahl, bei halben Werten von null weg, und lösen einen `RangeError` nur bei `NaN` oder unendlichen Eingaben aus. `roundTwips` und `roundEighthPoints` begrenzen negative Ergebnisse auf null.

**Umrechnungen aus realen Einheiten** runden auf die nächste ganze Zieleinheit: `twipsFromPoints`, `signedTwipsFromPoints`, `twipsFromPixels` (15 Twips pro CSS-Pixel), `signedTwipsFromPixels`, `twipsFromInches`, `halfPointsFromPoints`, `eighthPointsFromPoints`, `emuFromPixels`.

**Umrechnungen zurück** liefern einfache Zahlen: `twipsToPoints`, `twipsToPixels`, `twipsToInches`, `halfPointsToPoints`, `eighthPointsToPoints`, `emuToPixels`.

Außerhalb von `units.ts` gibt es keine Casts der Form `as Twips`. Verwenden Sie einen Konstruktor.

## Welches Modellfeld welches Brand trägt {#which-model-fields-carry-which-brand}

Die Feldnamen bleiben unverändert.

- **Absatz** (`Paragraph`, `ParagraphFormatting`, Style-Definitionen): `spacingBeforeTwips`, `spacingAfterTwips`, `firstLineTwips` und `hangingTwips` sind `Twips`; `lineSpacingTwips`, `indentLeftTwips`, `indentRightTwips`, `indentStartTwips` und `indentEndTwips` sind `SignedTwips`. `TabStop.posTwips` ist `SignedTwips`.
- **Nummerierungsebenen und Listenbeschriftungen**: `indentLeftTwips` ist `SignedTwips`; `hangingTwips` und `firstLineTwips` sind `Twips`.
- **Abschnitte**: `pageWidthTwips`, `pageHeightTwips`, `marginLeftTwips`, `marginRightTwips`, `headerDistanceTwips`, `footerDistanceTwips`, `gutterTwips` sowie die Spalten-`widthTwips`/`spacingTwips` sind `Twips`; `marginTopTwips` und `marginBottomTwips` sind `SignedTwips` (Word erlaubt negative obere und untere Ränder).
- **Tabellen**: `Table.grid` ist `Twips[]`; `Table.widthTwips`, `TableCell.widthTwips` und `TableRowProperties.heightTwips` sind `Twips`; `Table.indentTwips` ist `SignedTwips` (Word schreibt negative Tabelleneinzüge regelmäßig); `TableCellMargins.top/bottom/left/right` sind `SignedTwips`.
- **Rahmen**: `TableBorderSide.sizeEighthPoints` ist `EighthPoints`.
- **Textläufe**: `TextRun.characterSpacingTwips` ist `SignedTwips`.
- **Layout-Eingabe** (`ooxml-core/docx/layout`): Die Absatz-Felder `*Twips` von `LayoutParagraph` verwenden dieselben Brands wie das Modell.

### Felder, die bewusst nicht gebrandet sind {#fields-that-are-deliberately-not-branded}

- `TextRun.fontSize` steht in **Punkt** und darf Nachkommastellen haben (`10.5`). `HalfPoints` ist der auf der Festplatte gespeicherte Wert `w:sz`; der Parser rechnet ihn um. Verwenden Sie `halfPointsFromPoints`, wenn Sie eine Größe zurückschreiben.
- `TableBorderSide.spacePoints` sind ganze **Punkte** (`ST_PointMeasure`).
- `widthPx`, `heightPx`, `offsetXPx`, `offsetYPx` und jedes Layout-Feld `*Px` sind **CSS-Pixel bei 96 dpi**. Sie sind abgeleitete Werte im Gerätebereich und oft gebrochen, deshalb bleiben sie einfache Zahlen. Bildgrößen werden in Pixeln gespeichert (aus EMU abgeleitet) und nicht als `Emu`.
- `DocumentModel.page` (`width`, `height`, Ränder) steht in CSS-Pixeln; `SectionProperties` ist die maßgebliche Twip-Quelle.

## Modellwerte erstellen {#building-model-values}

Modellfelder erwarten gebrandete Zahlen. Code, der Modellobjekte erstellt oder bearbeitet, konstruiert sie deshalb:

```ts
import { twips, eighthPoints, roundSignedTwips, twipsFromPixels } from 'docx-core';
paragraph.spacingAfterTwips = twips(240);
table.grid = [twips(3000), twips(3000)];
side.sizeEighthPoints = eighthPoints(8);

// Computed values: round explicitly at the boundary
paragraph.indentLeftTwips = roundSignedTwips(dragPx * 15);
section.pageWidthTwips = twipsFromPixels(page.width);
```

Das Lesen ist davon nicht betroffen: Gebrandete Werte sind Zahlen, daher funktionieren `paragraph.spacingAfterTwips ?? 0` und `twipsToPixels(x)` wie gewohnt. Wenn Sie eine einfache Zahl in einen gebrandeten Platz einsetzen müssen, aus einer Quelle, der Sie vertrauen, dass sie ganzzahlig ist (zum Beispiel ein Wert, den Sie aus Ihrer eigenen JSON-Datei zurücklesen), übergeben Sie ihn über `twips()` oder `signedTwips()`, damit ein fehlerhafter Wert laut fehlschlägt. `twips(n)` lehnt negative Zahlen ab; verwenden Sie `signedTwips(n)` für vorzeichenbehaftete Werte.

## Parsen folgt der Vorzeichenregel des Schemas {#parsing-follows-the-schema-signedness}

Da das Modell vorzeichenbehaftete und vorzeichenlose Längen unterscheidet, wenden die Parser den Schematyp pro Attribut an. Negative Werte von `w:spacing/@before|@after`, `w:ind/@firstLine|@hanging`, der Seitenränder `w:pgMar/@left|@right|@header|@footer|@gutter`, des Tabellenrasters, der Breiten und der Zeilenhöhen werden abgelehnt (das Feld bleibt undefiniert). Dagegen bleiben negative Werte von `w:ind/@left|@right|@start|@end`, `w:spacing/@line`, `w:tab/@pos`, `w:pgMar/@top|@bottom`, `w:tblInd` und `w:spacing` (Zeichenabstand) erhalten. `validateModel` (der Validator vor dem Speichern) wendet dieselbe Einordnung an und lehnt Nicht-Ganzzahlen ab.

## Grenze der Editor-Attribute {#editor-attribute-boundary}

Die Attribute von ProseMirror-Knoten sind untypisiert. Der Editor wandelt sie an einer einzigen Stelle in das Modell um (`attr-units.ts` im `web-component`): Werte werden auf ganze Twips gerundet, und nicht numerische, nicht endliche oder (bei vorzeichenlosen Feldern) negative Werte werden verworfen, statt geschrieben zu werden.
