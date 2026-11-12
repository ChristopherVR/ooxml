# Funciones admitidas y garantía de conservación

Esta página es la lista escrita de funciones y la garantía de conservación que la
[política de estabilidad](../../../../docs/stability.md) exige a un producto antes de llamarlo Estable. Cubre
Word (`docx-core` y el editor). Indica qué está modelado y es editable, qué se transporta sin comprenderse
durante un guardado, y qué se rechaza, con el diagnóstico correspondiente. No afirma que Word y este editor
distribuyan un documento de la misma manera, ni que un archivo guardado sea el mismo que escribiría Word.

## Cómo funciona un guardado {#how-a-save-works}

Un archivo cargado con `loadDocx` conserva su paquete original. `save` vuelve a abrir ese paquete y cambia
solo las partes que alcanzan las ediciones:

- `word/document.xml` se analiza a un DOM y se vuelve a escribir desde él. El cuerpo se parchea bloque a
  bloque: un párrafo, tabla o sección que el modelo no cambió conserva su XML de origen, y el marcado que el
  modelo no contiene dentro de uno modificado (elementos desconocidos, atributos de extensión, contenido
  alternativo `mc:`, marcadores, envoltorios de cambios controlados) permanece donde estaba.
- Una parte solo se reescribe cuando el modelo cambió lo que contiene: una parte de encabezado, pie, nota al
  pie o nota final, los comentarios y sus partes de metadatos extendidos e de identidad, `word/numbering.xml`,
  las partes de estilos para el formato de estilo de párrafo editado, `docProps/core.xml`, las marcas de
  `word/settings.xml` que gestiona el editor, las partes de gráficos y SmartArt editadas, y el elemento
  `customXml` vinculado de un control con datos enlazados editado.
- `[Content_Types].xml` y las partes de relaciones solo cambian para registrar una parte o relación que añade el guardado.
- Cada otra entrada se vuelve a escribir tal como se leyó: mismo nombre, mismos bytes. Guardar un modelo que no
  cambió devuelve los bytes originales.

Un paquete creado por el editor (no cargado) no tiene origen que conservar y se escribe desde el modelo con los
estilos predeterminados integrados.

## Conservado tal cual {#kept-verbatim}

La prueba de ida y vuelta [`preservation.test.ts`](../../../../src/core/docx/preservation.test.ts) carga un
paquete que contiene estos elementos, edita texto, una celda de tabla y los comentarios, guarda, y verifica que
cada entrada que la edición no poseía es idéntica byte a byte (la lista de entradas también es la misma, salvo
las partes de comentarios creadas):

| Se conserva                                                                                                                                                                                                                                   | Notas                                                                                                                          |
| --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------ |
| Elementos `customXml/*`, sus propiedades y relaciones                                                                                                                                                                                         | Se reescriben solo para un control con datos enlazados que el usuario editó.                                                   |
| `word/glossary/document.xml` y sus partes (bloques de construcción)                                                                                                                                                                           | Se leen solo para mostrar el texto de marcador de posición de los controles de contenido.                                      |
| `word/vbaProject.bin` y el tipo de contenido con macros                                                                                                                                                                                       | Se conservan, nunca se ejecutan. La extensión del archivo guardado es responsabilidad del anfitrión.                           |
| `word/people.xml`                                                                                                                                                                                                                             | Consulte "Comentarios" más abajo.                                                                                              |
| `word/fontTable.xml` y las fuentes incrustadas (`word/fonts/*.odttf`)                                                                                                                                                                         | No se usan para el renderizado; se conservan con sus relaciones.                                                               |
| `word/settings.xml` (zoom, `w:extLst`, elementos de compatibilidad y protección)                                                                                                                                                              | Solo se parchean las marcas que gestiona el editor (control de cambios, encabezados pares e impares, guionado, ...).           |
| `word/theme/theme1.xml`, `word/webSettings.xml`, `docProps/custom.xml`, miniaturas                                                                                                                                                            | La edición del tema se rechaza (abajo).                                                                                        |
| Imágenes y otros medios que no se sustituyen                                                                                                                                                                                                  | Una imagen que el usuario sustituye recibe una parte nueva.                                                                    |
| Partes y relaciones desconocidas de cualquier tipo                                                                                                                                                                                            | Todo lo que contiene el paquete y para lo que el editor no tiene lector.                                                       |
| Dentro de `document.xml`: `mc:AlternateContent`, listas de extensiones (`w:extLst`), marcado `w14:`/`w15:`, `w:tblPrChange`, `w:trPrChange`, `w:tcPrChange`, marcadores, códigos de campo, propiedades de fragmento y de párrafo desconocidas | En párrafos y tablas que el usuario no reestructuró. Los párrafos editados conservan sus hijos `w:pPr` y `w:rPr` no modelados. |

## Comentarios {#comments}

Las ediciones de comentarios parchean los cuerpos y los metadatos originales en sus destinos de relación. Un
comentario sin cambios conserva sus párrafos, el formato de sus fragmentos y su marcado desconocido. Las ediciones
de texto conservan los fragmentos con estilo que sobreviven. El último ID de párrafo y el ID duradero mantienen
alineados los metadatos de respuestas, resolución y fecha UTC en `commentsExtended.xml`, `commentsIds.xml` y
`commentsExtensible.xml`. Los comentarios eliminados pierden sus anclas y metadatos; las nuevas respuestas reciben
referencias en el ámbito de su padre. Los identificadores opacos del editor se asignan de forma coherente a los
identificadores decimales del paquete, igual que sus anclas en el documento. La parte de personas se mantiene
tal cual. Esto no ofrece controles de redacción de comentarios con texto enriquecido.

La cobertura de regresión usa un fixture nativo de Word con formato en `comment-preservation.test.ts`.
Las comprobaciones de reapertura y guardado con Word 16 cubren resolver, editar, eliminar y añadir comentarios,
incluidas las respuestas y las ubicaciones de partes no estándar. Seis bindings de navegador comprueban las anclas
de respuestas guardadas.

## Modelado y editable {#modeled-and-editable}

Texto y formato de caracteres, formato de párrafo y estilos (uso y ediciones de formato por estilo), listas y
definiciones de numeración, secciones y configuración de página, encabezados, pies y notas al pie, campos
(resultados editables, Actualizar campos), comentarios y cambios controlados, controles de contenido y campos de
formulario, tablas (más abajo), vínculos y marcadores, imágenes, formas simples y cuadros de texto, SmartArt,
gráficos clásicos y ChartEx admitido, propiedades del documento, y las restricciones de edición que lee Word
(la protección de formulario se aplica; las demás restricciones se conservan pero no se aplican).

### Tablas {#tables}

Los comandos de filas y columnas, la combinación y división de celdas, y el diseño de celdas combinadas se editan
mediante el motor de estructura de tablas del núcleo (`ooxml-core/docx`, `table-structure.ts`) y se escriben con el
escritor que tiene en cuenta las combinaciones (`write-table-structure.ts`). El motor sigue a Word: sus resultados
para insertar fila, insertar columna, eliminar fila, eliminar columna, combinar y dividir se registraron con Word 16
en tablas pequeñas (`src/core/docx/__fixtures__/word-table-structure.json`, generado por
`src/core/scripts/docx/table-structure-word-probe.ps1`), y una prueba comprueba cada uno.

- Una fila o celda que sobrevive a una edición conserva su elemento XML, de modo que los hijos de `tblPr`, `trPr` y
  `tcPr` que el modelo no contiene, los atributos de extensión y los valores `w14:paraId` de fila se mantienen. Solo
  cambian `tblGrid`, `gridSpan`, `vMerge`, `tcW`, las celdas que crea una edición y las filas que añade o elimina.
- Las filas nuevas copian la fila de la que proceden (diseño de celdas, altura, repetir como encabezado); las celdas
  nuevas copian el formato de la celda contigua. Eliminar la primera fila de una combinación vertical traslada la
  combinación a la fila siguiente y vacía su texto, como hace Word.
- Las tablas anidadas se editan en su lugar: una edición de la tabla interior cambia solo el `w:tbl` interior.
- Una edición solo de texto en cualquier tabla nunca reescribe su estructura.

Se rechazan (el comando queda deshabilitado, y un cambio del modelo que el escritor no puede expresar lanza el
error de más abajo). Cada rechazo corresponde a una tabla que Word convierte en algo que una cuadrícula simple no
puede contener:

| Edición                                                                                                                                                                                                                                                                                                                    | Motivo                                                                                                                                                                                                                    |
| -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Insertar una columna a la izquierda de una celda combinada (una celda que abarca varias columnas), o a través de ella                                                                                                                                                                                                      | Word añade filas `gridAfter` y una celda nueva tan ancha como la celda combinada.                                                                                                                                         |
| Eliminar una columna que cubre solo en parte una celda combinada                                                                                                                                                                                                                                                           | Word elimina la celda combinada completa y deja la fila corta (`gridAfter`).                                                                                                                                              |
| Dividir una celda que abarca varias columnas en un número de columnas distinto, o en más columnas de las que abarca cuando abarca más de una                                                                                                                                                                               | Word vuelve a generar la cuadrícula de la tabla. Dividir en las columnas propias de la celda, y dividir una celda de una sola columna, sí funcionan.                                                                      |
| Dividir una combinación vertical en un número de filas distinto de su altura                                                                                                                                                                                                                                               | Word acepta solo divisores de la altura; solo se implementa la separación completa.                                                                                                                                       |
| Combinar un rango con una celda que sobresale de él, o con celdas que contienen tablas anidadas                                                                                                                                                                                                                            | El resultado perdería contenido o la forma de la combinación.                                                                                                                                                             |
| Cualquier edición estructural de una tabla con desplazamientos de fila (`gridBefore`, `gridAfter`), cambios controlados de fila o celda (`w:ins`, `w:del`, `trPrChange`, `tcPrChange`), `hMerge`, controles de contenido alrededor de filas o celdas, o vínculos, campos, marcadores u otro marcado en línea en sus celdas | La advertencia de carga "Las tablas combinadas, anidadas o complejas se pueden leer, pero su estructura de filas y columnas no se puede editar de forma segura." lo indica; el texto de sus celdas sigue siendo editable. |

`saveDocx` lanza `Cannot change the structure of a merged, nested, or complex table. The original
DOCX package remains unchanged.` ante un cambio estructural en una tabla así; el paquete no se modifica.

Las propiedades de nivel de tabla distintas de los márgenes que edita el diálogo Propiedades de tabla (ancho,
alineación, sangría, bordes, estilo, aspecto) se muestran y se conservan, pero no son editables: cambiarlas en el
modelo lanza `Cannot edit table grid widths, width, alignment, indent, borders, or style; ...`.

## Rechazado con un diagnóstico {#refused-with-a-diagnostic}

Guardar lanza una excepción, indicando el motivo y dejando el paquete de origen intacto, cuando el modelo:

- edita el catálogo de estilos de párrafo, de carácter o de tabla (`Editing the paragraph style catalog is not
  supported; source styles.xml is preserved unchanged.`, y las variantes de carácter y de tabla), salvo el formato por
  estilo que gestiona el editor;
- edita el tema (`Editing the theme catalog is not supported; source theme1.xml/settings.xml are preserved unchanged.`);
- repite un id de bloque o de párrafo (`Duplicate document block id` / `Duplicate document node id`);
- solicita una imagen, un vínculo, una forma o un gráfico que necesita una relación que el llamador no puede asignar;
- no supera la validación del modelo (`assertValidDocumentModel`: por ejemplo, un documento sin sección).

La carga está limitada a 50 MB por paquete, 5.000 entradas y 200 MB descomprimidos (`DOCX_ZIP_LIMITS`); un paquete que
supere un límite se rechaza antes de leer ninguna parte.

## Informado al cargar, no modificado por un guardado {#reported-at-load-not-changed-by-a-save}

`model.warnings` enumera las funciones que se cargan pero que no se renderizan ni editan como en Word, por ejemplo
objetos OLE incrustados (se conservan, se dibujan como marcador de posición), campos `HYPERLINK` complejos (se conservan
como códigos de campo), cambios controlados de estructura de tabla (se conservan en el XML, no en el modelo), viñetas
de imagen y parte de la numeración vinculada a estilos, y restricciones de edición que se conservan pero no se aplican.

## Límites conocidos de la garantía {#known-limits-of-the-guarantee}

- **El marcado en línea dentro de las celdas de tabla** bloquea las ediciones estructurales de esa tabla (arriba), pero no
  las ediciones de texto.
- **El diseño no es información conservada.** Las pistas de salto de página en caché (`w:lastRenderedPageBreak`) de los
  párrafos editados se descartan; Word las recalcula.
- El contenido de una parte para la que el editor no tiene lector se conserva, no se valida. Una parte dañada sigue dañada.
