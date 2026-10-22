---
title: Das PptxData-Modell
description: Die Struktur des analysierten PptxData-Modells - Folien, Design, Master und Metadaten - dazu die vollständige diskriminierte Union PptxElement, Typwächter und EMU-Einheitenumrechnungen.
---

# Das PptxData-Modell

`handler.load()` liefert ein einzelnes `PptxData`-Objekt: die vollständig analysierte Darstellung einer Präsentation im Arbeitsspeicher. Bearbeiten heißt, dieses Objekt zu verändern; Speichern serialisiert es zurück in ein `.pptx`-Archiv.

## `PptxData`

Das Modell der obersten Ebene. Die am häufigsten verwendeten Felder sind hier aufgeführt; die Schnittstelle trägt außerdem umfangreiche Metadaten für einen verlustfreien Round-Trip.

```ts
interface PptxData {
	slides: PptxSlide[]; // the slides, in order
	width: number; // presentation width in pixels (approx)
	height: number; // presentation height in pixels (approx)
	widthEmu?: number; // slide width in EMU (round-trip)
	heightEmu?: number; // slide height in EMU (round-trip)
	slideSizeType?: string; // e.g. "screen4x3", "screen16x9", "custom"

	theme?: PptxTheme; // full parsed theme (colours, fonts, name)
	themeColorMap?: Record<string, string>; // colour scheme key → hex
	slideMasters?: PptxSlideMaster[]; // masters (each with its layouts)

	sections?: PptxSection[]; // ordered presentation sections
	customShows?: PptxCustomShow[]; // named custom slide shows
	embeddedFonts?: PptxEmbeddedFont[];

	// document metadata
	coreProperties?: PptxCoreProperties; // docProps/core.xml
	appProperties?: PptxAppProperties; // docProps/app.xml
	customProperties?: PptxCustomProperty[]; // docProps/custom.xml

	// flags
	isPasswordProtected?: boolean;
	hasMacros?: boolean; // is a .pptm
	hasDigitalSignatures?: boolean;
	conformance?: 'strict' | 'transitional'; // original OOXML conformance class

	warnings?: PptxCompatibilityWarning[]; // unsupported-feature notices
}
```

::: info Pixel und EMU
`width` / `height` sind ungefähre Pixelwerte für bequeme Layoutrechnung; `widthEmu` / `heightEmu` bewahren die exakten Quellwerte für den Round-Trip. Siehe [Einheiten](#einheiten-emu-und-pixel) unten.
:::

## `PptxSlide`

Jeder Eintrag in `data.slides`. Das Array `elements` enthält den erstellten Inhalt.

```ts
interface PptxSlide {
	id: string;
	rId: string; // relationship ID
	slideNumber: number;
	name?: string; // optional author-supplied name
	layoutPath?: string; // the layout this slide is based on
	layoutName?: string;
	hidden?: boolean; // hidden slides are skipped in presentation mode
	sectionName?: string;
	sectionId?: string;
	elements: PptxElement[]; // the slide's content
	backgroundColor?: string;
	backgroundImage?: string; // base64 data URL
	backgroundGradient?: string; // CSS gradient string
	backgroundPattern?: PptxSlideBackgroundPattern;
	transition?: PptxSlideTransition;
	animations?: PptxElementAnimation[];
}
```

Eine Folie löst nicht angegebene Formatierung über ihr Layout und ihren Master auf - siehe [Folien, Layouts und Master](#folien-layouts-und-master) unten.

## Die Elementbasis und Mixins

Jede Elementvariante erweitert `PptxElementBase` (definiert in `src/core/pptx/core/types/element-base.ts`), die Identität, Geometrie und Round-Trip-Daten trägt:

```ts
interface PptxElementBase {
	id: string; // synthetic positional id assigned by the loader
	shapeId?: string; // native OOXML id from p:cNvPr/@id (animation target key)
	name?: string; // cNvPr/@name (used for morph "!!" matching)

	x: number; // pixels, converted from EMU at parse time
	y: number;
	width: number;
	height: number;
	rotation?: number; // degrees
	skewX?: number; // degrees
	skewY?: number;
	flipHorizontal?: boolean;
	flipVertical?: boolean;

	hidden?: boolean;
	opacity?: number; // 0-1
	locks?: PptxShapeLocks; // p:cNvSpPr/a:spLocks
	actionClick?: PptxAction; // a:hlinkClick on p:cNvPr
	actionHover?: PptxAction; // a:hlinkHover on p:cNvPr

	rawXml?: XmlObject; // preserved source XML
	extLstXml?: XmlObject[]; // unrecognised <a:ext> extensions, kept verbatim
}
```

Zwei Mixin-Schnittstellen ergänzen Fähigkeiten. Textführende Elemente fügen `PptxTextProperties` hinzu (`text`, `textStyle`, formatierte `textSegments`, `promptText`, Verkettung verknüpfter Textfelder); Formen, Verbinder und Bilder fügen `PptxShapeProperties` hinzu (`shapeType`, `shapeStyle`, `shapeAdjustments`, `adjustmentHandles`). Zum Beispiel:

```ts
interface TextPptxElement extends PptxElementBase, PptxTextProperties, PptxShapeProperties {
	type: 'text';
}

interface ImagePptxElement
	extends PptxElementBase, PptxShapeProperties, PptxCustomPathProperties, PptxImageProperties {
	type: 'image';
}
```

::: tip Round-Trip-Felder
`rawXml` und `extLstXml` sorgen dafür, dass das Modell verlustfrei bleibt: Markup, das das typisierte Modell (noch) nicht versteht, bleibt am Element erhalten und wird beim Speichern wieder ausgegeben. Entfernen Sie diese Felder nicht, wenn Sie speichern wollen.
:::

## Die `PptxElement`-Union

`slide.elements` ist ein Array der diskriminierten Union `PptxElement`: **16 Varianten**, definiert in `src/core/pptx/core/types/elements.ts`. Grenzen Sie über das Feld `type` ein, um auf variantenspezifische Eigenschaften zuzugreifen.

| `type`-String   | Schnittstelle            | Beschreibung                                                                                      |
| --------------- | ------------------------ | ------------------------------------------------------------------------------------------------- |
| `"text"`        | `TextPptxElement`        | Ein Textfeld - ein Rechteck mit Text, meist ohne sichtbare Füllung oder Kontur.                   |
| `"shape"`       | `ShapePptxElement`       | Eine Form mit Voreinstellungs- oder Freiformgeometrie; kann auch Text enthalten.                  |
| `"connector"`   | `ConnectorPptxElement`   | Eine gerade, gewinkelte oder gebogene Linie, die Formen verbindet, mit optionalen Pfeilspitzen.   |
| `"image"`       | `ImagePptxElement`       | Ein Rasterbild aus einem `<p:pic>`-Knoten.                                                        |
| `"picture"`     | `PicturePptxElement`     | Funktional identisch mit `image`, aus semantischer Klarheit über die Diskriminante unterschieden. |
| `"table"`       | `TablePptxElement`       | Eine über einen `<p:graphicFrame>` eingebettete Tabelle; Zellen in `tableData`.                   |
| `"chart"`       | `ChartPptxElement`       | Ein über einen `<p:graphicFrame>` eingebettetes Diagramm; Daten in `chartData`.                   |
| `"smartArt"`    | `SmartArtPptxElement`    | Ein in positionierte Formen zerlegtes SmartArt-Diagramm; Daten in `smartArtData`.                 |
| `"ole"`         | `OlePptxElement`         | Ein OLE-Objekt (eingebettetes Excel/Word/PDF/Visio/MathType) mit Vorschaubild.                    |
| `"media"`       | `MediaPptxElement`       | Ein Audio- oder Videoelement, das auf eine Datei im Archiv verweist.                              |
| `"group"`       | `GroupPptxElement`       | Ein Container, dessen `children` die Transformation der Gruppe erben.                             |
| `"ink"`         | `InkPptxElement`         | Freihand-Striche von Stift/Maus als SVG-Pfade gespeichert, mit optionalem Druck.                  |
| `"contentPart"` | `ContentPartPptxElement` | Ein Inhaltsteil `mc:AlternateContent`, typischerweise moderne Stift-/Textmarker-Tinte.            |
| `"zoom"`        | `ZoomPptxElement`        | Ein Folienzoom- oder Abschnittszoom-Objekt, das zu einer Zielfolie navigiert.                     |
| `"model3d"`     | `Model3DPptxElement`     | Ein über `p16:model3D` eingebettetes 3D-Modell (`.glb`/`.gltf`) mit Posterbild.                   |
| `"unknown"`     | `UnknownPptxElement`     | Ein Element, das der Parser nicht erkennt; für den Round-Trip erhalten.                           |

Graphic-Frame-Varianten (`table`, `chart`, `smartArt`) tragen außerdem ein Array `extensionXml` mit nicht erkannten `a:graphicData/a:extLst`-Erweiterungen, die für den Round-Trip unverändert erhalten bleiben.

Einige dieser Elementtypen, wie die Engine sie rendert (Diagramme als Inline-SVG, Tabellen als HTML, Formen und Verbinder als SVG-Geometrie):

![Ein Balkendiagramm-Element als Inline-SVG gerendert](/docs-shots/chart-slide.jpg)

![Ein Tabellenelement als HTML-Tabelle gerendert](/docs-shots/table-slide.jpg)

![Voreingestellte Formen und Verbinder als SVG-Geometrie gerendert](/docs-shots/shapes-slide.jpg)

### Teilmengen-Aliase

Für Funktionssignaturen, die eine Teilmenge akzeptieren, exportiert das Core-Paket Hilfsaliase:

```ts
type PptxElementWithText = TextPptxElement | ShapePptxElement | ConnectorPptxElement;
type PptxImageLikeElement = ImagePptxElement | PicturePptxElement;
// plus PptxElementWithShapeStyle for everything carrying shapeStyle
```

## Typwächter

`src/core/pptx/core/types/type-guards.ts` exportiert Laufzeitwächter, mit denen Sie eingrenzen können, ohne `element.type === ...`-Vergleiche von Hand zu schreiben:

| Wächter                  | Grenzt ein auf                                           |
| ------------------------ | -------------------------------------------------------- |
| `isTextElement(el)`      | `TextPptxElement`                                        |
| `isShapeElement(el)`     | `ShapePptxElement`                                       |
| `isConnectorElement(el)` | `ConnectorPptxElement`                                   |
| `isImageLikeElement(el)` | `PptxImageLikeElement` (`image` oder `picture`)          |
| `isInkElement(el)`       | `InkPptxElement`                                         |
| `isZoomElement(el)`      | `ZoomPptxElement`                                        |
| `hasTextProperties(el)`  | `PptxElementWithText` (kann Text tragen)                 |
| `hasShapeProperties(el)` | `PptxElementWithShapeStyle` (kann einen Formstil tragen) |

```ts
import { isImageLikeElement, hasTextProperties } from 'pptx-viewer-core';

for (const el of slide.elements) {
	if (isImageLikeElement(el)) console.log(el.imagePath);
	if (hasTextProperties(el)) console.log(el.text);
}
```

## Einheiten: EMU und Pixel

Das native Koordinatensystem von PowerPoint ist die **English Metric Unit** (EMU):

| Konstante                      | Wert                | Wo                                            |
| ------------------------------ | ------------------- | --------------------------------------------- |
| `EMU_PER_PX` / `EMU_PER_PIXEL` | `9525` (bei 96 DPI) | `core/constants.ts` und das SDK-Modul `units` |
| `EMU_PER_INCH`                 | `914400`            | SDK-Modul `units`                             |
| `EMU_PER_POINT`                | `12700`             | SDK-Modul `units`                             |

Positionen und Größen von Elementen im Modell sind **Pixel**, beim Parsen aus EMU umgerechnet (und gerundet) und beim Speichern zurückgerechnet. Das Builder-SDK exportiert außerdem Umrechnungshelfer:

```ts
import {
	inches,
	cm,
	mm,
	pt, // to pixels
	emuToPixels,
	pixelsToEmu, // EMU <-> px
	inchesToEmu,
	cmToEmu, // to EMU (e.g. for slide dimensions)
	SlideSizes,
} from 'pptx-viewer-core';

inches(1); // => 96 (px)
pt(12); // => 16 (px)
pixelsToEmu(96); // => 914400
SlideSizes.WIDESCREEN_16_9; // => { width: 12192000, height: 6858000 } (EMU)
```

`SlideSizes` liefert die üblichen Foliengrößen in EMU: `WIDESCREEN_16_9` (der moderne Standard), `STANDARD_4_3`, `WIDESCREEN_16_10`, `A4_LANDSCAPE`, `A4_PORTRAIT`, `LETTER_LANDSCAPE` und `LETTER_PORTRAIT`.

## Folien, Layouts und Master

Die Formatierung einer Präsentation ist hierarchisch. Jede Folie verweist auf ein **Layout** (über `slide.layoutPath`), jedes Layout gehört zu einem **Master**, und jeder Master verweist auf ein **Design**:

```
PptxSlide ── layoutPath ──> PptxSlideLayout ──> PptxSlideMaster ── themePath ──> PptxTheme
```

`data.slideMasters` stellt diesen Baum bereit:

```ts
interface PptxSlideMaster {
	path: string; // e.g. "ppt/slideMasters/slideMaster1.xml"
	name?: string;
	themePath?: string; // theme this master references
	layoutPaths?: string[]; // layouts belonging to this master
	layouts?: PptxSlideLayout[]; // parsed layout objects
	elements?: PptxElement[]; // shapes drawn on the master itself
	placeholders?: Array<{ type: string; idx?: string }>;
	txStyles?: PptxMasterTextStyles; // title/body/other text defaults
	clrMap?: Record<string, string>; // scheme-colour alias map (bg1, tx1, ...)
	backgroundColor?: string;
	backgroundImage?: string;
}

interface PptxSlideLayout {
	path: string;
	name?: string;
	elements?: PptxElement[];
	placeholders?: Array<{ type: string; idx?: string }>;
	clrMapOverride?: Record<string, string>; // p:clrMapOvr
	backgroundColor?: string;
	backgroundImage?: string;
}
```

Wenn ein Folienelement eine Eigenschaft nicht angibt, löst die Engine sie die Kette hinauf auf: Element, dann der passende Platzhalter im Layout, dann der Master, dann das Design. `themeColorMap` in `PptxData` ist die bereits aufgelöste Schemafarben-Zuordnung für den Standardmaster. Die Auflösungsregeln stehen unter [Architektur](/de/guide/architecture#auflosungskette-fur-designs).

## Iterieren und eingrenzen

```ts
import { PptxHandler } from 'pptx-viewer-core';
import type { PptxElement } from 'pptx-viewer-core';

const handler = new PptxHandler();
const data = await handler.load(buffer);

function summarize(elements: PptxElement[]) {
	for (const element of elements) {
		switch (element.type) {
			case 'text':
				console.log('text:', element.text);
				break;
			case 'image':
			case 'picture':
				console.log('image:', element.imagePath);
				break;
			case 'table':
				console.log('table rows:', element.tableData?.rows.length ?? 0);
				break;
			case 'chart':
				console.log('chart:', element.chartData?.chartType);
				break;
			case 'group':
				summarize(element.children); // recurse into the group
				break;
			default:
				console.log('element:', element.type);
		}
	}
}

for (const slide of data.slides) {
	summarize(slide.elements);
}
```

## Weiterführend

- [Architektur](/de/guide/architecture) - wie die Engine aufgebaut ist und wie die Designauflösung funktioniert.
- [Core-Paketübersicht](/core/) - APIs für Handler, Builder und Konverter.
- [OpenXML-Konformität](/architecture/openxml-conformance) - Round-Trip von Strict und Transitional.
