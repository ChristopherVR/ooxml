---
title: Le modèle PptxData
description: La forme du modèle PptxData analysé - diapositives, thème, masques et métadonnées - ainsi que l'union discriminée PptxElement complète, les gardes de type et les conversions d'unités EMU.
---

# Le modèle PptxData

`handler.load()` renvoie un seul objet `PptxData` : la représentation en mémoire, entièrement analysée, d'une présentation. Éditer revient à modifier cet objet ; sauvegarder le sérialise de nouveau en archive `.pptx`.

## `PptxData`

Le modèle de plus haut niveau. Les champs les plus utilisés sont listés ici ; l'interface porte aussi de nombreuses métadonnées pour un aller-retour sans perte.

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

::: info Pixels et EMU
`width` / `height` sont des pixels approximatifs pour faciliter les calculs de mise en page ; `widthEmu` / `heightEmu` conservent les valeurs exactes de la source pour l'aller-retour. Voir [Unités](#unites-emu-et-pixels) ci-dessous.
:::

## `PptxSlide`

Chaque entrée de `data.slides`. Le tableau `elements` contient le contenu créé.

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

Une diapositive résout son style non spécifié via sa mise en page et son masque ; voir [Diapositives, mises en page et masques](#diapositives-mises-en-page-et-masques) ci-dessous.

## La base des éléments et les mixins

Chaque variante d'élément étend `PptxElementBase` (défini dans `src/core/pptx/core/types/element-base.ts`), qui porte l'identité, la géométrie et les données d'aller-retour :

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

Deux interfaces de mixin ajoutent des capacités. Les éléments qui portent du texte ajoutent `PptxTextProperties` (`text`, `textStyle`, `textSegments` riches, `promptText`, chaînage de zones de texte liées) ; les formes, connecteurs et images ajoutent `PptxShapeProperties` (`shapeType`, `shapeStyle`, `shapeAdjustments`, `adjustmentHandles`). Par exemple :

```ts
interface TextPptxElement extends PptxElementBase, PptxTextProperties, PptxShapeProperties {
	type: 'text';
}

interface ImagePptxElement
	extends PptxElementBase, PptxShapeProperties, PptxCustomPathProperties, PptxImageProperties {
	type: 'image';
}
```

::: tip Champs d'aller-retour
`rawXml` et `extLstXml` sont ce qui garde le modèle sans perte : le balisage que le modèle typé ne comprend pas (encore) est conservé sur l'élément et réémis à la sauvegarde. Ne supprimez pas ces champs si vous comptez sauvegarder.
:::

## L'union `PptxElement`

`slide.elements` est un tableau de l'union discriminée `PptxElement` : **16 variantes**, définies dans `src/core/pptx/core/types/elements.ts`. Restreignez sur le champ `type` pour accéder aux propriétés propres à une variante.

| Chaîne `type`   | Interface                | Description                                                                                            |
| --------------- | ------------------------ | ------------------------------------------------------------------------------------------------------ |
| `"text"`        | `TextPptxElement`        | Une zone de texte : un rectangle contenant du texte, généralement sans remplissage ni contour visible. |
| `"shape"`       | `ShapePptxElement`       | Une forme à géométrie prédéfinie ou libre ; peut aussi contenir du texte.                              |
| `"connector"`   | `ConnectorPptxElement`   | Une ligne droite, coudée ou courbe reliant des formes, avec pointes de flèche optionnelles.            |
| `"image"`       | `ImagePptxElement`       | Une image raster issue d'un nœud `<p:pic>`.                                                            |
| `"picture"`     | `PicturePptxElement`     | Fonctionnellement identique à `image`, distinguée par son discriminant pour la clarté sémantique.      |
| `"table"`       | `TablePptxElement`       | Un tableau incorporé via un `<p:graphicFrame>` ; cellules dans `tableData`.                            |
| `"chart"`       | `ChartPptxElement`       | Un graphique incorporé via un `<p:graphicFrame>` ; données dans `chartData`.                           |
| `"smartArt"`    | `SmartArtPptxElement`    | Un diagramme SmartArt décomposé en formes positionnées ; données dans `smartArtData`.                  |
| `"ole"`         | `OlePptxElement`         | Un objet OLE (Excel/Word/PDF/Visio/MathType incorporé) avec une image d'aperçu.                        |
| `"media"`       | `MediaPptxElement`       | Un élément audio ou vidéo référençant un fichier de l'archive.                                         |
| `"group"`       | `GroupPptxElement`       | Un conteneur dont les `children` héritent de la transformation du groupe.                              |
| `"ink"`         | `InkPptxElement`         | Des traits à main levée (stylet/souris) stockés en tracés SVG, avec pression optionnelle.              |
| `"contentPart"` | `ContentPartPptxElement` | Une partie de contenu `mc:AlternateContent`, typiquement l'encre moderne de stylo/surligneur.          |
| `"zoom"`        | `ZoomPptxElement`        | Un objet Zoom de diapositive ou de section qui navigue vers une diapositive cible.                     |
| `"model3d"`     | `Model3DPptxElement`     | Un modèle 3D (`.glb`/`.gltf`) incorporé via `p16:model3D`, avec une image d'affiche.                   |
| `"unknown"`     | `UnknownPptxElement`     | Un élément que l'analyseur ne reconnaît pas ; préservé pour l'aller-retour.                            |

Les variantes de cadre graphique (`table`, `chart`, `smartArt`) portent aussi un tableau `extensionXml` d'extensions `a:graphicData/a:extLst` non reconnues, préservées telles quelles pour l'aller-retour.

Quelques-uns de ces types d'éléments tels que le moteur les affiche (graphiques en SVG en ligne, tableaux en HTML, formes et connecteurs en géométrie SVG) :

![Un élément graphique à barres affiché en SVG en ligne](/docs-shots/chart-slide.jpg)

![Un élément tableau affiché en tableau HTML](/docs-shots/table-slide.jpg)

![Formes prédéfinies et connecteurs affichés en géométrie SVG](/docs-shots/shapes-slide.jpg)

### Alias de sous-ensembles

Pour les signatures de fonctions qui acceptent un sous-ensemble, le package Core exporte des alias utilitaires :

```ts
type PptxElementWithText = TextPptxElement | ShapePptxElement | ConnectorPptxElement;
type PptxImageLikeElement = ImagePptxElement | PicturePptxElement;
// plus PptxElementWithShapeStyle for everything carrying shapeStyle
```

## Gardes de type

`src/core/pptx/core/types/type-guards.ts` exporte des gardes d'exécution pour restreindre le type sans écrire à la main des comparaisons `element.type === ...` :

| Garde                    | Restreint à                                                 |
| ------------------------ | ----------------------------------------------------------- |
| `isTextElement(el)`      | `TextPptxElement`                                           |
| `isShapeElement(el)`     | `ShapePptxElement`                                          |
| `isConnectorElement(el)` | `ConnectorPptxElement`                                      |
| `isImageLikeElement(el)` | `PptxImageLikeElement` (`image` ou `picture`)               |
| `isInkElement(el)`       | `InkPptxElement`                                            |
| `isZoomElement(el)`      | `ZoomPptxElement`                                           |
| `hasTextProperties(el)`  | `PptxElementWithText` (peut porter du texte)                |
| `hasShapeProperties(el)` | `PptxElementWithShapeStyle` (peut porter un style de forme) |

```ts
import { isImageLikeElement, hasTextProperties } from 'pptx-viewer-core';

for (const el of slide.elements) {
	if (isImageLikeElement(el)) console.log(el.imagePath);
	if (hasTextProperties(el)) console.log(el.text);
}
```

## Unités : EMU et pixels

Le système de coordonnées natif de PowerPoint est l'**English Metric Unit** (EMU) :

| Constante                      | Valeur            | Où                                           |
| ------------------------------ | ----------------- | -------------------------------------------- |
| `EMU_PER_PX` / `EMU_PER_PIXEL` | `9525` (à 96 DPI) | `core/constants.ts` et le module SDK `units` |
| `EMU_PER_INCH`                 | `914400`          | module SDK `units`                           |
| `EMU_PER_POINT`                | `12700`           | module SDK `units`                           |

Les positions et tailles des éléments dans le modèle sont en **pixels**, converties (et arrondies) depuis les EMU à l'analyse et reconverties à la sauvegarde. Le SDK de construction exporte aussi des utilitaires de conversion :

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

`SlideSizes` fournit les dimensions standard de présentation en EMU : `WIDESCREEN_16_9` (la valeur moderne par défaut), `STANDARD_4_3`, `WIDESCREEN_16_10`, `A4_LANDSCAPE`, `A4_PORTRAIT`, `LETTER_LANDSCAPE` et `LETTER_PORTRAIT`.

## Diapositives, mises en page et masques

Le style d'une présentation est hiérarchique. Chaque diapositive référence une **mise en page** (via `slide.layoutPath`), chaque mise en page appartient à un **masque**, et chaque masque référence un **thème** :

```
PptxSlide ── layoutPath ──> PptxSlideLayout ──> PptxSlideMaster ── themePath ──> PptxTheme
```

`data.slideMasters` expose cet arbre :

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

Lorsqu'un élément de diapositive laisse une propriété non spécifiée, le moteur la résout en remontant la chaîne : l'élément, puis l'espace réservé correspondant de la mise en page, puis le masque, puis le thème. `themeColorMap` sur `PptxData` est la table de couleurs de schéma déjà résolue pour le masque par défaut. Les règles de résolution sont décrites dans [Architecture](/fr/guide/architecture#chaine-de-resolution-des-themes).

## Parcourir et restreindre

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

## Lectures connexes

- [Architecture](/fr/guide/architecture) - structure du moteur et fonctionnement de la résolution des thèmes.
- [Aperçu du package Core](/core/) - API du handler, du builder et du convertisseur.
- [Conformité OpenXML](/architecture/openxml-conformance) - aller-retour Strict et Transitional.
