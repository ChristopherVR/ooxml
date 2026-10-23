---
title: Arquitectura
description: Cómo está estructurado pptx-viewer - las canalizaciones de carga y guardado, el runtime compuesto por mixins, la resolución del tema, el motor de geometría y la capa de renderizado compartida que alimenta cada enlace de framework.
---

# Arquitectura

`pptx-viewer` es un sistema por capas. Los enlaces de framework (React, Vue 3, Angular, Svelte y un enlace vanilla JS sobre DOM puro) se encargan de la interfaz; una capa de renderizado compartida aporta la lógica independiente del framework; y el motor core se ocupa de todo lo relacionado con analizar, editar y guardar archivos de PowerPoint. Cada capa depende únicamente de la inferior.

## Descripción general

```
+---------------------------------------------------------------+
|                      Framework bindings                       |
|   React    |   Vue 3   |  Angular  |  Svelte 5  |  Vanilla JS |
| pptx-react | pptx-vue  | pptx-ang. | pptx-svelte| pptx-vanilla|
+------------------------------+--------------------------------+
                               |
                               v
+---------------------------------------------------------------+
|             Shared rendering layer (ooxml-ui/pptx)            |
|     geometry, styles, gradients, charts, connectors, text     |
+------------------------------+--------------------------------+
                               |
                               v
+---------------------------------------------------------------+
|                Core engine (pptx-viewer-core)                 |
|  PptxHandler (public API)                                     |
|    -> Runtime (parsing, serialization, theme resolution)      |
|         -> Types, Geometry, Colour, Builders, Converter       |
+---------------------------------------------------------------+
```

## Enlaces de framework

Los paquetes de enlace son capas de presentación delgadas. Consumen datos de renderizado precalculados de `ooxml-ui/pptx` y los traducen a plantillas específicas de cada framework (JSX, SFC de Vue, componentes de Angular, runes de Svelte 5 o llamadas DOM simples en el enlace vanilla). Las diapositivas se renderizan como HTML/SVG escalado con transformaciones CSS, lo que da texto nítido a cualquier zoom, accesibilidad nativa e interactividad DOM completa.

Cada enlace expone un punto de entrada de visualizador/editor de nivel superior que orquesta el estado, la edición, la carga, la exportación y el modo de presentación con el idioma propio de su framework: hooks en React, composables en Vue, servicios en Angular, runes en Svelte y una función fábrica simple con una API de instancia imperativa en vanilla JS. Como todos consumen la misma capa compartida, el resultado del renderizado es idéntico en los cinco.

## Capa de renderizado compartida (`ooxml-ui/pptx`)

El renderizado compartido vive en `src/ui/src/pptx/` en la raíz del repositorio y se publica mediante `ooxml-ui/pptx`. Las operaciones de documento sin DOM viven en `src/core/pptx/editor/`, y los esquemas de herramientas en `src/core/pptx/automation/schemas/`. El renderizador incluye:

- **Enrutamiento de conectores**: un enrutador A\* sobre un grafo de obstáculos (`connector-router-astar.ts`, `connector-router-graph.ts`, ahora en el área compartida `ooxml-core/geometry`) más la construcción y el reenrutado de trayectos.
- **Matemáticas de gráficos**: rangos de ejes, posicionamiento de categorías, constructores de trazados cartesianos/polares, estadísticas de cajas y bigotes, composición combinada/bursátil (la familia de módulos `chart-*`).
- **Motores de animación y transformación (morph)**: construcción y reproducción de líneas de tiempo (`animation-timeline-*.ts`), además de la correspondencia de transiciones morph y la interpolación de geometría (`morph-matching.ts`, `morph-geometry-interp.ts`, `morph-text.ts`).
- **Resolución de texto y estilo**: numeración de viñetas, estilos de relleno/degradado, filtros de efectos de imagen, estilos de salto de línea kinsoku.
- **Matemáticas**: conversores de OMML a MathML y de LaTeX a OMML (`omml-to-mathml.ts`, `latex-to-omml.ts`, ahora en el área compartida `ooxml-core/math`).
- **Comportamiento del editor**: portapapeles, guías de alineación, copiar formato, buscar y reemplazar, sincronización de colaboración y presencia.

Los directorios hermanos cubren `export/`, `i18n/`, `loader/`, `theme/` y el renderizador opcional `smartart-3d/`.

### Controladores sin interfaz

La lógica del editor con estado se escribe una vez como controlador sin interfaz en `render/controller/`: un objeto simple con `getSnapshot()`, `subscribe(listener)` y acciones tipadas. El historial de deshacer y rehacer (profundidad predeterminada de 100 pasos), el autoguardado de recuperación (un único motor que también guarda el texto en línea sin confirmar), la oferta de recuperación, el ciclo de vida de la sala de colaboración y el panel de IA son controladores, igual que los gestos del lienzo (`createGestureController`). El despachador de teclas (`dispatchEditorKey`), el puente de IA (`createPptxAiBridge`), los destinos de pulsación en grupos y los pasos de la presentación son funciones puras junto a ellos. Cada enlace conecta su propio estado reactivo a un adaptador de almacén genérico (`useController` en React y Vue, `controllerState` en Svelte, `controllerSignal` en Angular, `subscribe` en vanilla), de modo que una corrección en un controlador llega a los cinco enlaces.

Las secciones Efectos, Línea, Tamaño y posición, y Cuadro de texto del inspector son descriptores de campos en `render/inspector/`, que dibuja un único renderizador genérico por enlace. La hoja de datos del gráfico, el panel de formato del gráfico, los diálogos de insertar gráfico y SmartArt, el panel de texto de SmartArt y las galerías de diseño de gráficos son elementos compartidos de `ooxml-ui` que todos los enlaces montan.

::: info Paquetes públicos compartidos
Cada enlace declara `ooxml-ui` como dependencia de ejecución e importa sus entradas de PowerPoint. Angular usa la misma API pública sin incluir una copia del código fuente. El paquete privado `pptx-viewer-shared` sigue siendo una fachada de compatibilidad y no posee ninguna implementación.
:::

## Motor core (`pptx-viewer-core`)

El paquete core es totalmente independiente del framework. Se ejecuta en cualquier entorno JavaScript: navegador, Node.js, Web Worker o función sin servidor. Su punto de entrada público es `PptxHandler`.

> **Dónde vive el código.** El motor es el área `pptx` del paquete público compartido `ooxml-core` (`src/core/pptx/`, subrutas `/pptx`, `/pptx/converter`, `/pptx/cli`, `/pptx/signature-node`); la carpeta del visualizador conserva los adaptadores de framework. `pptx-viewer-core` es un paquete delgado que reexporta esas subrutas con una API pública sin cambios, por lo que toda ruta de abajo que empiece por `src/core/pptx/` es relativa a la raíz del monorepo. El motor sigue usando su propio modelo de objetos `fast-xml-parser` y se compila con indicadores de TypeScript relajados; unificar su modelo XML con el área compartida `xml` y endurecer los indicadores son los próximos pasos de la migración. Las primitivas de color, la geometría de formas predefinidas y las constantes EMU se comparten con las demás áreas de Office (`color`, `geometry`, `units`).

### La fachada y el runtime compuesto por mixins

El motor se construye como tres capas de superficie decreciente:

```
PptxHandler                    static factories (create / createBlank)
  └─ PptxHandlerCore           thin facade: load / save / export / encryption
       └─ IPptxHandlerRuntime  the actual engine, assembled from many mixin modules
```

- **`PptxHandler`** (`src/core/pptx/core/PptxHandler.ts`) añade los puntos de entrada de construcción estáticos `create()` / `createBlank()`.
- **`PptxHandlerCore`** delega todo el análisis pesado, la serialización y la manipulación de XML en un `IPptxHandlerRuntime` inyectado. El runtime es reemplazable mediante dependencias del constructor (`runtime` o `runtimeFactory`), que es como las pruebas y los hosts alternativos intercambian implementaciones.
- **`PptxHandlerRuntime`** no es una sola clase en un solo archivo. Se compone de muchos **módulos específicos** en `src/core/pptx/core/core/runtime/`, cada uno llamado `PptxHandlerRuntime<Concern>.ts` y que se ocupa exactamente de una responsabilidad: `PptxHandlerRuntimeChartParsing.ts`, `PptxHandlerRuntimeThemeLoading.ts`, `PptxHandlerRuntimeSaveElementWriter.ts`, `PptxHandlerRuntimeSmartArtParsing.ts`, etc.

Cada módulo declara una clase que extiende la clase exportada por el módulo anterior, formando una cadena de herencia lineal que va apilando capacidad sobre capacidad:

```ts
// PptxHandlerRuntimeLoadPipeline.ts
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeLoadSession';

export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	// adds the load-pipeline capability on top of everything below it
}
```

Las nuevas capacidades del motor (incluidos los nuevos tipos de elemento) se añaden como eslabones nuevos de esta cadena en lugar de hacer crecer los archivos existentes. Los colaboradores transversales (advertencias de compatibilidad, fábricas de XML, constructores de tipos de contenido) se definen tras interfaces de servicio `I*` y se inyectan mediante una fábrica de dependencias, lo que mantiene cada mixin comprobable de forma aislada.

### La canalización de carga

`handler.load(arrayBuffer)` recorre estas etapas:

```
ArrayBuffer
  │  container sniff: an OLE compound file is either a legacy binary
  │  .ppt (converted to an in-memory .pptx package) or an encrypted
  │  OOXML package (decrypted with the supplied password)
  │  ZIP signature check on whatever comes out
  ▼
JSZip.loadAsync                 in-memory archive
  │  zip-bomb guard: 500 MiB uncompressed budget (configurable),
  │  hard cap of 65,536 entries
  ▼
fast-xml-parser                 XML parts -> JS object trees
  │  Strict OOXML detection: strict namespace URIs are transparently
  │  normalized to Transitional for all subsequent parses
  ▼
Theme / master / layout resolution
  │  colour maps, font schemes, format schemes, placeholder styles
  ▼
Per-slide shape-tree (spTree) parsing
  │  each <p:sp>, <p:pic>, <p:graphicFrame>, ... becomes a typed PptxElement
  ▼
PptxLoadDataBuilder             assembles the final model
  ▼
PptxData
```

1. **Primero las comprobaciones del contenedor.** Un archivo compuesto OLE no es necesariamente un error: es como llegan tanto un `.ppt` binario heredado como un paquete OOXML cifrado. Una presentación de PowerPoint 97-2003 (reconocida por su flujo `PowerPoint Document`) se convierte en un paquete `.pptx` equivalente en memoria mediante `core/ppt/` y se carga desde ahí, de modo que el resto de la canalización nunca sabe que la presentación era binaria. En caso contrario, el contenedor se trata como cifrado y lanza `EncryptedFileError` salvo que se proporcione una contraseña para descifrarlo; un `.ppt` protegido con contraseña lanza `EncryptedPptError`, porque el cifrado RC4 heredado no se descifra. Una entrada que no sea ni un ZIP ni un archivo compuesto OLE se rechaza con un error claro. Los archivos sobredimensionados lanzan `ZipBombError` antes de que comience cualquier análisis.
2. **El análisis XML** usa `fast-xml-parser`; cada parte se convierte en un árbol de objetos simple (el tipo `XmlObject`) con atributos prefijados por `@_`.
3. **La resolución del tema** carga el tema de cada patrón, el mapa de colores (`p:clrMap`), el esquema de fuentes y el esquema de formato, de modo que el análisis de diapositivas pueda resolver los colores de esquema y las referencias de estilo (véase más abajo).
4. **El análisis de elementos** convierte el árbol de formas de cada diapositiva en la [unión discriminada `PptxElement`](/es/guide/data-model#la-union-pptxelement), convirtiendo las coordenadas EMU en píxeles y conservando el XML sin procesar de las construcciones que el modelo tipado no cubre.
5. **`PptxLoadDataBuilder`** ensambla todo lo demás: secciones, presentaciones personalizadas, fuentes incrustadas, patrones de notas y de documentos, etiquetas, autores de comentarios, propiedades del documento, la miniatura y las advertencias de compatibilidad.

### La canalización de guardado

`handler.save(slides, options?)` invierte el proceso (`PptxHandlerRuntimeSavePipeline.ts`):

```
PptxSlide[]
  │  resolve conformance class ('preserve' | 'strict' | 'transitional')
  ▼
Reconcile presentation slide list      order, additions, deletions, rels
  ▼
Serialize each slide to OpenXML        elements -> <p:spTree>, embed new media
  ▼
Rebuild [Content_Types].xml            slide overrides + media defaults
  ▼
Comments, masters, layouts             typed mutations applied; untouched
  │                                    parts pass through verbatim
  ▼
Optional Strict conversion             remap namespaces if target is strict
  ▼
JSZip -> Uint8Array                    a valid .pptx / .ppsx / .pptm
```

Dos propiedades de esta canalización importan para la fidelidad:

- **Paso directo por defecto.** Las partes que no editó (patrones, diseños, notas, extensiones desconocidas, marcado de proveedores) se trasladan del archivo original byte a byte o se vuelven a emitir a partir de sus árboles de análisis conservados. Guardar es una reescritura dirigida, no una regeneración completa.
- **Salida consciente de la conformidad.** El paquete guardado coincide por defecto con la clase de conformidad OOXML del archivo cargado, y puede forzarse a Strict o Transitional. Consulte [Conformidad con OpenXML](/architecture/openxml-conformance) para ver exactamente qué se reasigna.

### Cadena de resolución del tema

Los estilos de PowerPoint se resuelven mediante una cadena de herencia, y el motor la reproduce exactamente:

```
Element  ->  Placeholder  ->  Layout  ->  Master  ->  Theme
```

- Las propiedades explícitas de un elemento siempre prevalecen.
- Un elemento marcador de posición (un título, cuerpo, pie, ...) hereda los valores por defecto de texto y posición del marcador correspondiente de su **diseño**, que a su vez hereda del **patrón** (`p:txStyles` para los valores por defecto de texto de título/cuerpo/otros).
- Las referencias de estilo (`a:fillRef`, `a:lnRef`, `a:effectRef`, `a:fontRef`) se resuelven por índice en el esquema de formato del **tema**, sustituyendo el color de marcador del tema por el color de la referencia.
- Los colores de esquema (`accent1`, `bg1`, `tx1`, ...) se resuelven mediante el `p:clrMap` del patrón (y cualquier `p:clrMapOvr` del diseño) en el esquema de colores del tema.

La cadena la implementan mixins de runtime dedicados: `PptxHandlerRuntimeThemeLoading`, `...ThemeProcessing`, `...ThemeFormatScheme`, `...ThemeRefResolution`, `...ThemeOverrides`, y el trío de marcadores de posición `...PlaceholderLookup`, `...PlaceholderDefaults`, `...PlaceholderStyles`. Se admiten varios patrones, cada uno con su propio mapa de colores y esquema de formato.

### Motor de geometría

`src/core/pptx/core/geometry/` (más las tablas de formas predefinidas compartidas de `src/core/geometry/`) convierte la geometría DrawingML en trazados renderizables:

- **Formas predefinidas**: definiciones del catálogo de formas predefinidas de ECMA-376, agrupadas por familia (`preset-shape-definitions-arrows.ts`, `-flowchart.ts`, `-action-buttons.ts`, `-callouts`, ...), cada una expresada con las fórmulas de guía de la especificación.
- **Evaluación de fórmulas de guía** (`guide-formula-eval.ts` y archivos afines): implementa el lenguaje de fórmulas de ECMA-376 (`*/`, `+-`, `pin`, `at2`, `cos`, ...) para que la geometría de las formas responda correctamente a los valores de ajuste, los tiradores amarillos en forma de rombo que se pueden arrastrar en PowerPoint.
- **Geometría personalizada** (`custom-geometry-parser.ts`, `freeform-builder.ts`): analiza los comandos de trazado `a:custGeom` y construye formas libres.
- **Trazados de recorte** (`preset-clip-paths-core.ts`, `-extended.ts`): producen trazados de recorte CSS/SVG para que el contenido HTML (imágenes, texto) pueda recortarse a cualquier forma predefinida.
- **Geometría de conectores** (`connector-geometry.ts`): conectores angulares y curvos, puntas de flecha y gestión de volteos. (El **enrutamiento** dinámico de conectores alrededor de obstáculos vive en el enrutador A\* de la capa compartida.)

### Conversor

`src/core/pptx/converter/` implementa la conversión de PPTX a Markdown con un patrón de registro: cada tipo de elemento tiene un procesador (`shape-element-processor`, `table-element-processor`, `ole-element-processor`, ...) registrado contra su discriminante `type`, y `PptxMarkdownConverter` despacha por elemento. El mismo directorio alberga el exportador SVG y el conversor de OMML a LaTeX que se usa para las ecuaciones.

## Decisiones de diseño clave

| Decisión                                  | Justificación                                                                                                                                                                                                                                                                       |
| ----------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Renderizado basado en CSS** (no Canvas) | Texto nítido a cualquier zoom, accesibilidad nativa, interactividad DOM y estilos CSS estándar.                                                                                                                                                                                     |
| **Motor compuesto por mixins**            | Muchos módulos de runtime pequeños y de una sola responsabilidad mantienen cada capacidad aislada y comprobable. Las nuevas capacidades (incluidos los nuevos tipos de elemento) se añaden como mixins nuevos, no como archivos más grandes.                                        |
| **Unión discriminada para los elementos** | TypeScript estrecha al tipo de elemento correcto mediante el campo `type`, con seguridad de tipos total y sin conversiones.                                                                                                                                                         |
| **Cadena de resolución del tema**         | Elemento, marcador de posición, diseño, patrón y tema reproduce la herencia de estilos propia de PowerPoint.                                                                                                                                                                        |
| **Unidades EMU internamente**             | PowerPoint usa English Metric Units (914.400 EMU por pulgada; 9.525 EMU por píxel a 96 DPI). Los elementos analizados exponen valores en píxeles para facilitar los cálculos de diseño; los valores EMU exactos se conservan donde la fidelidad del viaje de ida y vuelta lo exige. |
| **Guardado con paso directo**             | Solo se reescriben las partes editadas; todo lo demás hace el viaje de ida y vuelta sin cambios, de modo que el marcado desconocido y las extensiones de proveedores sobreviven.                                                                                                    |
| **Lógica compartida, enlaces delgados**   | Toda la lógica del visualizador independiente del framework vive una sola vez en `ooxml-ui/pptx`; los enlaces son solo capas de vista, lo que mantiene los cinco frameworks en paridad.                                                                                             |

## Lecturas relacionadas

- [El modelo PptxData](/es/guide/data-model) - la forma completa de una presentación analizada, los tipos de elemento y las unidades.
- [Conformidad con OpenXML](/architecture/openxml-conformance) - tratamiento de Strict frente a Transitional y el contrato de conformidad.
- [Resumen del paquete core](/core/) - la referencia de la API pública.
