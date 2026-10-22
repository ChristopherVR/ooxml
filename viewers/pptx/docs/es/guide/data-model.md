---
title: El modelo PptxData
description: La forma del modelo PptxData analizado (diapositivas, tema, patrones y metadatos), además de la unión discriminada completa PptxElement, las guardas de tipo y las conversiones de unidades EMU.
---

# El modelo PptxData

`handler.load()` devuelve un único objeto `PptxData`: la representación completa, analizada y en memoria de una presentación. Editar significa mutar este objeto; guardar lo serializa de nuevo en un archivo `.pptx`.

## `PptxData`

El modelo de nivel superior. Aquí se enumeran los campos más usados; la interfaz también lleva abundantes metadatos para un viaje de ida y vuelta sin pérdidas.

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

::: info Píxeles frente a EMU
`width` / `height` son píxeles aproximados para facilitar los cálculos de diseño; `widthEmu` / `heightEmu` conservan los valores exactos del origen para el viaje de ida y vuelta. Consulte [Unidades](#unidades-emu-y-pixeles) más abajo.
:::

## `PptxSlide`

Cada entrada de `data.slides`. El arreglo `elements` contiene el contenido creado.

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

Una diapositiva resuelve el estilo no especificado a través de su diseño y su patrón; consulte [Diapositivas, diseños y patrones](#diapositivas-disenos-y-patrones) más abajo.

## La base de elemento y los mixins

Cada variante de elemento extiende `PptxElementBase` (definido en `src/core/pptx/core/types/element-base.ts`), que lleva identidad, geometría y datos de ida y vuelta:

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

Dos interfaces mixin añaden capacidades. Los elementos con texto añaden `PptxTextProperties` (`text`, `textStyle`, `textSegments` enriquecidos, `promptText`, encadenamiento de cuadros de texto vinculados); las formas, los conectores y las imágenes añaden `PptxShapeProperties` (`shapeType`, `shapeStyle`, `shapeAdjustments`, `adjustmentHandles`). Por ejemplo:

```ts
interface TextPptxElement extends PptxElementBase, PptxTextProperties, PptxShapeProperties {
	type: 'text';
}

interface ImagePptxElement
	extends PptxElementBase, PptxShapeProperties, PptxCustomPathProperties, PptxImageProperties {
	type: 'image';
}
```

::: tip Campos de ida y vuelta
`rawXml` y `extLstXml` son la forma en que el modelo se mantiene sin pérdidas: el marcado que el modelo tipado no entiende (todavía) se conserva en el elemento y se vuelve a emitir al guardar. No elimine estos campos si piensa guardar.
:::

## La unión `PptxElement`

`slide.elements` es un arreglo de la unión discriminada `PptxElement`: **16 variantes**, definidas en `src/core/pptx/core/types/elements.ts`. Estreche el tipo con el campo `type` para acceder a las propiedades específicas de cada variante.

| Cadena `type`   | Interfaz                 | Descripción                                                                                      |
| --------------- | ------------------------ | ------------------------------------------------------------------------------------------------ |
| `"text"`        | `TextPptxElement`        | Un cuadro de texto: un rectángulo que contiene texto, normalmente sin relleno ni trazo visibles. |
| `"shape"`       | `ShapePptxElement`       | Una forma con geometría predefinida o libre; también puede contener texto.                       |
| `"connector"`   | `ConnectorPptxElement`   | Una línea recta, angular o curva que une formas, con puntas de flecha opcionales.                |
| `"image"`       | `ImagePptxElement`       | Una imagen raster de un nodo `<p:pic>`.                                                          |
| `"picture"`     | `PicturePptxElement`     | Funcionalmente idéntica a `image`, distinguida por el discriminante por claridad semántica.      |
| `"table"`       | `TablePptxElement`       | Una tabla incrustada mediante un `<p:graphicFrame>`; celdas en `tableData`.                      |
| `"chart"`       | `ChartPptxElement`       | Un gráfico incrustado mediante un `<p:graphicFrame>`; datos en `chartData`.                      |
| `"smartArt"`    | `SmartArtPptxElement`    | Un diagrama SmartArt descompuesto en formas posicionadas; datos en `smartArtData`.               |
| `"ole"`         | `OlePptxElement`         | Un objeto OLE (Excel/Word/PDF/Visio/MathType incrustado) con una imagen de vista previa.         |
| `"media"`       | `MediaPptxElement`       | Un elemento de audio o video que referencia un archivo del paquete.                              |
| `"group"`       | `GroupPptxElement`       | Un contenedor cuyos `children` heredan la transformación del grupo.                              |
| `"ink"`         | `InkPptxElement`         | Trazos a mano alzada de lápiz o ratón almacenados como trazados SVG, con presión opcional.       |
| `"contentPart"` | `ContentPartPptxElement` | Una parte de contenido `mc:AlternateContent`, normalmente tinta moderna de lápiz o resaltador.   |
| `"zoom"`        | `ZoomPptxElement`        | Un objeto de Zoom de diapositiva o de sección que navega a una diapositiva de destino.           |
| `"model3d"`     | `Model3DPptxElement`     | Un modelo 3D (`.glb`/`.gltf`) incrustado mediante `p16:model3D`, con una imagen de póster.       |
| `"unknown"`     | `UnknownPptxElement`     | Un elemento que el analizador no reconoce; se conserva para el viaje de ida y vuelta.            |

Las variantes de marco gráfico (`table`, `chart`, `smartArt`) también llevan un arreglo `extensionXml` con las extensiones `a:graphicData/a:extLst` no reconocidas, conservadas literalmente para el viaje de ida y vuelta.

Algunos de estos tipos de elemento tal como los renderiza el motor (gráficos como SVG en línea, tablas como HTML, formas y conectores como geometría SVG):

![A bar chart element rendered as inline SVG](/docs-shots/chart-slide.jpg)

![A table element rendered as an HTML table](/docs-shots/table-slide.jpg)

![Preset shapes and connectors rendered as SVG geometry](/docs-shots/shapes-slide.jpg)

### Alias de subconjuntos

Para las firmas de funciones que aceptan un subconjunto, el paquete core exporta alias auxiliares:

```ts
type PptxElementWithText = TextPptxElement | ShapePptxElement | ConnectorPptxElement;
type PptxImageLikeElement = ImagePptxElement | PicturePptxElement;
// plus PptxElementWithShapeStyle for everything carrying shapeStyle
```

## Guardas de tipo

`src/core/pptx/core/types/type-guards.ts` exporta guardas en tiempo de ejecución para que pueda estrechar el tipo sin escribir a mano comparaciones `element.type === ...`:

| Guarda                   | Estrecha a                                                 |
| ------------------------ | ---------------------------------------------------------- |
| `isTextElement(el)`      | `TextPptxElement`                                          |
| `isShapeElement(el)`     | `ShapePptxElement`                                         |
| `isConnectorElement(el)` | `ConnectorPptxElement`                                     |
| `isImageLikeElement(el)` | `PptxImageLikeElement` (`image` o `picture`)               |
| `isInkElement(el)`       | `InkPptxElement`                                           |
| `isZoomElement(el)`      | `ZoomPptxElement`                                          |
| `hasTextProperties(el)`  | `PptxElementWithText` (puede llevar texto)                 |
| `hasShapeProperties(el)` | `PptxElementWithShapeStyle` (puede llevar estilo de forma) |

```ts
import { isImageLikeElement, hasTextProperties } from 'pptx-viewer-core';

for (const el of slide.elements) {
	if (isImageLikeElement(el)) console.log(el.imagePath);
	if (hasTextProperties(el)) console.log(el.text);
}
```

## Unidades: EMU y píxeles

El sistema de coordenadas nativo de PowerPoint es la **English Metric Unit** (EMU):

| Constante                      | Valor             | Dónde                                           |
| ------------------------------ | ----------------- | ----------------------------------------------- |
| `EMU_PER_PX` / `EMU_PER_PIXEL` | `9525` (a 96 DPI) | `core/constants.ts` y el módulo `units` del SDK |
| `EMU_PER_INCH`                 | `914400`          | Módulo `units` del SDK                          |
| `EMU_PER_POINT`                | `12700`           | Módulo `units` del SDK                          |

Las posiciones y tamaños de los elementos en el modelo son **píxeles**, convertidos (y redondeados) desde EMU al analizar y reconvertidos al guardar. El SDK de construcción también exporta auxiliares de conversión:

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

`SlideSizes` ofrece las dimensiones estándar de presentación en EMU: `WIDESCREEN_16_9` (el valor moderno por defecto), `STANDARD_4_3`, `WIDESCREEN_16_10`, `A4_LANDSCAPE`, `A4_PORTRAIT`, `LETTER_LANDSCAPE` y `LETTER_PORTRAIT`.

## Diapositivas, diseños y patrones

El estilo de una presentación es jerárquico. Cada diapositiva referencia un **diseño** (mediante `slide.layoutPath`), cada diseño pertenece a un **patrón** y cada patrón referencia un **tema**:

```
PptxSlide ── layoutPath ──> PptxSlideLayout ──> PptxSlideMaster ── themePath ──> PptxTheme
```

`data.slideMasters` expone este árbol:

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

Cuando un elemento de diapositiva deja una propiedad sin especificar, el motor la resuelve subiendo por la cadena: el elemento, luego el marcador de posición correspondiente del diseño, luego el patrón y por último el tema. `themeColorMap` en `PptxData` es la tabla de colores de esquema ya resuelta del patrón por defecto. Las reglas de resolución se describen en [Arquitectura](/es/guide/architecture#cadena-de-resolucion-del-tema).

## Iterar y estrechar tipos

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

## Lecturas relacionadas

- [Arquitectura](/es/guide/architecture) - cómo está estructurado el motor y cómo funciona la resolución del tema.
- [Resumen del paquete core](/core/) - APIs del handler, del constructor y del conversor.
- [Conformidad con OpenXML](/architecture/openxml-conformance) - viaje de ida y vuelta Strict frente a Transitional.
