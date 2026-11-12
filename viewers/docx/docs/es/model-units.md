# Unidades de medida en el modelo de documento

`DocumentModel` almacena las medidas en las unidades que usa WordprocessingML en el disco. Desde que se
incorporaron las marcas de unidad, esas unidades forman parte de los tipos de TypeScript: pasar un tamaño de
fuente donde se espera una longitud, o un valor de píxeles sin redondear donde se requieren twips enteros,
es un error de compilación en lugar de un archivo corrupto.

Las marcas existen solo en tiempo de compilación. Un `Twips` es un número de JavaScript normal en tiempo de
ejecución, y sigue siendo asignable a `number`, de modo que la aritmética, la comparación y la serialización
JSON no cambian.

## Marcas {#brands}

Todas las marcas se exportan desde `docx-core`.

| Tipo           | Unidad                                 | Tipo del esquema        | Constructor (estricto) | Constructor de redondeo |
| -------------- | -------------------------------------- | ----------------------- | ---------------------- | ----------------------- |
| `Twips`        | 1/20 de punto, no negativo             | `ST_TwipsMeasure`       | `twips(n)`             | `roundTwips(n)`         |
| `SignedTwips`  | 1/20 de punto, puede ser negativo      | `ST_SignedTwipsMeasure` | `signedTwips(n)`       | `roundSignedTwips(n)`   |
| `HalfPoints`   | 1/2 de punto (tamaño de fuente `w:sz`) | `ST_HpsMeasure`         | `halfPoints(n)`        | `roundHalfPoints(n)`    |
| `EighthPoints` | 1/8 de punto (ancho de borde)          | `ST_EighthPointMeasure` | `eighthPoints(n)`      | `roundEighthPoints(n)`  |
| `Emu`          | 1/914400 de pulgada (DrawingML)        | `ST_Coordinate`         | `emu(n)`               | `roundEmu(n)`           |

`Twips` es un subtipo de `SignedTwips`: una medida sin signo puede asignarse a un campo con signo, pero nunca
al revés.

**Los constructores estrictos** lanzan `RangeError` salvo que el valor ya sea un entero seguro y completo
(y, para `twips`, no negativo). Los analizadores y las pruebas los usan.

**Los constructores de redondeo** sirven para valores calculados (arrastres de la regla, cálculos de zoom,
conversiones a píxeles). Redondean la mitad alejándose del cero y solo lanzan `RangeError` para entradas
`NaN` o infinitas. `roundTwips` y `roundEighthPoints` fijan a cero los resultados negativos.

**Las conversiones desde unidades del mundo real** redondean al entero más cercano de la unidad de destino:
`twipsFromPoints`, `signedTwipsFromPoints`, `twipsFromPixels` (15 twips por píxel CSS),
`signedTwipsFromPixels`, `twipsFromInches`, `halfPointsFromPoints`, `eighthPointsFromPoints`,
`emuFromPixels`.

**Las conversiones de salida** devuelven números normales: `twipsToPoints`, `twipsToPixels`,
`twipsToInches`, `halfPointsToPoints`, `eighthPointsToPoints`, `emuToPixels`.

Fuera de `units.ts` no hay conversiones del tipo `as Twips`; use un constructor.

## Qué campos del modelo llevan cada marca {#which-model-fields-carry-which-brand}

Los nombres de los campos no cambian.

- **Párrafo** (`Paragraph`, `ParagraphFormatting`, definiciones de estilo): `spacingBeforeTwips`,
  `spacingAfterTwips`, `firstLineTwips`, `hangingTwips` son `Twips`; `lineSpacingTwips`,
  `indentLeftTwips`, `indentRightTwips`, `indentStartTwips`, `indentEndTwips` son `SignedTwips`.
  `TabStop.posTwips` es `SignedTwips`.
- **Niveles de numeración y etiquetas de lista**: `indentLeftTwips` es `SignedTwips`; `hangingTwips` y
  `firstLineTwips` son `Twips`.
- **Secciones**: `pageWidthTwips`, `pageHeightTwips`, `marginLeftTwips`, `marginRightTwips`,
  `headerDistanceTwips`, `footerDistanceTwips`, `gutterTwips` y los campos `widthTwips`/`spacingTwips`
  de las columnas son `Twips`; `marginTopTwips` y `marginBottomTwips` son `SignedTwips` (Word permite
  márgenes superiores e inferiores negativos).
- **Tablas**: `Table.grid` es `Twips[]`; `Table.widthTwips`, `TableCell.widthTwips` y
  `TableRowProperties.heightTwips` son `Twips`; `Table.indentTwips` es `SignedTwips` (Word escribe
  sangrías de tabla negativas con frecuencia); `TableCellMargins.top/bottom/left/right` son `SignedTwips`.
- **Bordes**: `TableBorderSide.sizeEighthPoints` es `EighthPoints`.
- **Fragmentos de texto**: `TextRun.characterSpacingTwips` es `SignedTwips`.
- **Entrada de diseño** (`ooxml-core/docx/layout`): los campos `*Twips` del párrafo en `LayoutParagraph`
  usan las mismas marcas que el modelo.

### Campos que deliberadamente no llevan marca {#fields-that-are-deliberately-not-branded}

- `TextRun.fontSize` está en **puntos** y puede ser fraccionario (`10.5`). `HalfPoints` es el valor
  `w:sz` tal como está en el disco; el analizador lo convierte. Use `halfPointsFromPoints` al escribir
  un tamaño de nuevo.
- `TableBorderSide.spacePoints` está en **puntos** enteros (`ST_PointMeasure`).
- `widthPx`, `heightPx`, `offsetXPx`, `offsetYPx` y todos los campos `*Px` del diseño son **píxeles CSS
  a 96 ppp**. Son valores derivados del espacio del dispositivo y con frecuencia fraccionarios, por lo que
  siguen siendo números normales. Las dimensiones de imagen se almacenan en píxeles (derivados de EMU) en
  lugar de como `Emu`.
- `DocumentModel.page` (`width`, `height`, márgenes) está en píxeles CSS; `SectionProperties` es la fuente
  autorizada de twips.

## Construir valores del modelo {#building-model-values}

Los campos del modelo aceptan números con marca, de modo que el código que crea o edita objetos del modelo
los construye:

```ts
import { twips, eighthPoints, roundSignedTwips, twipsFromPixels } from 'docx-core';
paragraph.spacingAfterTwips = twips(240);
table.grid = [twips(3000), twips(3000)];
side.sizeEighthPoints = eighthPoints(8);

// Computed values: round explicitly at the boundary
paragraph.indentLeftTwips = roundSignedTwips(dragPx * 15);
section.pageWidthTwips = twipsFromPixels(page.width);
```

La lectura no cambia: los valores con marca son números, así que `paragraph.spacingAfterTwips ?? 0` y
`twipsToPixels(x)` funcionan como siempre. Cuando necesite un número normal en un campo con marca desde una
fuente en la que confía que sea entera (por ejemplo, un valor leído de su propio JSON), páselo por `twips()`
o `signedTwips()` para que un valor incorrecto falle de forma explícita. `twips(n)` rechaza números negativos;
use `signedTwips(n)` para valores con signo.

## El análisis sigue la signedness del esquema {#parsing-follows-the-schema-signedness}

Como el modelo ahora distingue las longitudes con y sin signo, los analizadores aplican el tipo del esquema
atributo por atributo. Se rechazan (el campo queda sin definir) los valores negativos de
`w:spacing/@before|@after`, `w:ind/@firstLine|@hanging`, `w:pgMar/@left|@right|@header|@footer|@gutter`,
la cuadrícula de tablas, los anchos y las alturas de fila; en cambio se conservan los negativos de
`w:ind/@left|@right|@start|@end`, `w:spacing/@line`, `w:tab/@pos`, `w:pgMar/@top|@bottom`, `w:tblInd` y
`w:spacing` (espaciado de caracteres). `validateModel` (el validador previo al guardado) aplica la misma
clasificación y rechaza los valores no enteros.

## Frontera de atributos del editor {#editor-attribute-boundary}

Los atributos de nodo de ProseMirror no tienen tipo. El editor los convierte en el modelo en un único punto
(`attr-units.ts` en `web-component`): los valores se redondean a twips enteros, y los valores no numéricos,
no finitos o (para los campos sin signo) negativos se descartan en lugar de escribirse.
