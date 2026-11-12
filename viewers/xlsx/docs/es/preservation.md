# Funciones admitidas y garantía de conservación

Esta página es la lista escrita que [la política de estabilidad](https://github.com/ChristopherVR/ooxml/blob/main/docs/stability.md)
exige a un producto Stable. Dice qué modela y edita el editor de Excel, qué conserva intacto a través de una edición y un guardado, y qué descarta o rechaza (con el mensaje que muestra).
[Funciones y limitaciones](/es/features) describe el comportamiento visible para el usuario, y la
[lista de funciones](/functions) el motor de cálculo. Nada de aquí afirma que el editor sea igual a Microsoft Excel ni que un archivo guardado sea idéntico al original: la afirmación es más acotada, como se explica a continuación.

## La garantía {#the-guarantee}

El contenido que el editor no modela (elementos desconocidos, listas de extensiones, partes, relaciones,
XML personalizado y similares) sobrevive a abrir un libro, editarlo y guardarlo, o el editor informa de que no puede conservarlo. Una pérdida silenciosa es un error y se corrige en una versión de parche.

"Sobrevive" significa que el XML o los bytes vuelven sin cambios, en una posición válida. No significa que el paquete completo sea idéntico: las partes que el editor regenera a partir de su modelo (listadas más abajo) se reescriben, y Excel puede volver a reescribirlas en su próximo guardado.

## Lo que se modela y se edita {#what-is-modelled-and-edited}

Estos elementos se leen en el modelo, se editan mediante la sesión de edición y se escriben de nuevo desde el modelo.
Los atributos y los hijos dentro de ellos que el modelo no contiene se conservan donde la tabla lo indica.

| Área             | Modelado y escrito desde el modelo                                                                                                                                                                                                                                                            |
| ---------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Celdas           | Valores, fórmulas (las fórmulas compartidas se expanden), fórmulas de matriz y de matriz dinámica con sus metadatos, fórmulas de tabla de datos, cadenas de texto enriquecido, errores, fechas en ambos sistemas de fechas.                                                                   |
| Filas y columnas | Alturas, anchos, ocultas, niveles de esquema, contraídas, formatos de fila y de columna, indicaciones de fila gruesa superior e inferior.                                                                                                                                                     |
| Formatos         | Fuentes, rellenos, bordes, formatos de número, alineación, protección, estilos de celda con nombre, formatos diferenciales.                                                                                                                                                                   |
| Hoja             | Combinaciones, paneles inmovilizados, vistas, zoom, líneas de cuadrícula, color de la pestaña, estado (visible, oculta, muy oculta), protección con hashes de contraseña, configuración de página, opciones de impresión, área de impresión, encabezados y pies de página.                    |
| Reglas y listas  | Formato condicional (los tipos de regla admitidos), validación de datos, hipervínculos, filtros automáticos, tablas con sus filtros y su fila de totales, nombres definidos.                                                                                                                  |
| Anotaciones      | Notas, comentarios en hilo con sus personas.                                                                                                                                                                                                                                                  |
| Dibujos          | Gráficos (a través del modelo de gráficos compartido; un gráfico que el modelo no puede contener se conserva tal como se leyó), imágenes, formas que entiende, SmartArt (el dibujo en caché se regenera, los datos de origen se conservan), minigráficos, segmentaciones y escalas de tiempo. |
| Análisis         | Tablas dinámicas de rango de hoja que el editor admite, tablas de datos de hipótesis, escenarios del Administrador de escenarios.                                                                                                                                                             |
| Libro            | Orden y nombres de las hojas, hoja activa, nombres definidos (con los atributos de los nombres de macro y de función), configuración de cálculo, protección del libro, sistema de fechas, paleta y fuentes del tema.                                                                          |
| Propiedades      | Propiedades del documento principales, extendidas y personalizadas.                                                                                                                                                                                                                           |

## Lo que se conserva sin cambios a través de una edición y un guardado {#what-is-kept-unchanged-through-an-edit-and-a-save}

Las ediciones de comentarios en hilo conservan los autores, las fechas de creación, las identidades de los mensajes, los registros de cuenta de persona, el estado de resuelto y el marcado de comentario desconocido. Las menciones importadas siguen el texto insertado o eliminado alrededor de sus etiquetas sin cambios; cambiar una etiqueta elimina esa mención. Copiar un hilo asigna identidades nuevas de mensaje y de mención, pero conserva las personas a las que hace referencia. El editor puede resolver y reabrir hilos, pero no puede crear menciones @ nuevas.

El escritor copia una parte del paquete de origen byte a byte (con sus relaciones y cada parte a la que llegan) cuando nada en el modelo la regenera. La fuente de esto es `PackageWriter.carry` en `src/core/xlsx/write/package-writer.ts`, controlado por `saveXlsx`
(`src/core/xlsx/write/save.ts`) y `writeWorksheet` (`src/core/xlsx/write/worksheet.ts`).

| Conservado                                                                      | Cómo                                                                                                                                                                                                                                                                                                                                                                            |
| ------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Proyecto VBA (`xl/vbaProject.bin`) de un `.xlsm`                                | Se conserva con su relación y el tipo de contenido habilitado para macros. Las macros nunca se ejecutan. Guardar el libro como `.xlsx` elimina el proyecto, porque ese formato no puede contenerlo.                                                                                                                                                                             |
| XML personalizado (`customXml/*`) y sus partes de propiedades                   | Se conserva con la relación del libro y las relaciones entre las partes.                                                                                                                                                                                                                                                                                                        |
| Vínculos externos (`xl/externalLinks/*`)                                        | Se conservan con sus destinos de relación (externos); el elemento `externalReferences` permanece en `workbook.xml`. Los vínculos no se actualizan (se informa al cargar).                                                                                                                                                                                                       |
| Conexiones, tablas de consulta y otras partes del libro                         | Cualquier relación del libro que el escritor no regenera se conserva con su parte de destino, sea del tipo que sea.                                                                                                                                                                                                                                                             |
| Configuración de impresora, propiedades de control y propiedades personalizadas | Las relaciones de hoja que el escritor no regenera se conservan con sus partes de destino (`printerSettings`, `ctrlProps`, `customProperty`, objetos incrustados, imágenes).                                                                                                                                                                                                    |
| Elementos de hoja que el modelo no representa                                   | `sheetCalcPr`, `phoneticPr`, `rowBreaks`, `colBreaks`, `customProperties`, `cellWatches`, `ignoredErrors`, `smartTags`, `oleObjects`, `controls`, `picture`, `webPublishItems`, `extLst` y los demás del esquema se conservan como XML y se escriben de nuevo en el orden del esquema. Las entradas de extensión desconocidas dentro de `extLst` se mantienen tal como estaban. |
| Atributos de origen de elementos de hoja modelados                              | `sheetPr`, `sheetViews`, `sheetFormatPr`, `sheetProtection`, `autoFilter`, `pageMargins`, `pageSetup` y `headerFooter` conservan los atributos que el modelo no contiene cuando la parte modelada no cambió.                                                                                                                                                                    |
| Elementos del libro que el modelo no representa                                 | `fileVersion`, `fileSharing`, `functionGroups`, `externalReferences`, `oleSize`, `customWorkbookViews`, `smartTagPr`, `smartTagTypes`, `webPublishing`, `fileRecoveryPr`, `webPublishObjects`, `extLst`, cada ventana de `bookViews` y los atributos no modelados de `workbookPr`, `workbookView` y `calcPr`.                                                                   |
| Hojas de gráfico, hojas de diálogo, hojas de macros                             | Se conservan completas (se muestran como hojas vacías y se informa al cargar); sus referencias siguen los cambios de nombre y las eliminaciones de hojas.                                                                                                                                                                                                                       |
| Tablas dinámicas no admitidas, cachés de tablas dinámicas                       | Se conservan con sus relaciones; las definiciones de caché siguen los cambios de nombre de hojas.                                                                                                                                                                                                                                                                               |
| Dibujos y gráficos que el editor no tocó                                        | Las partes de dibujo, gráfico, estilo, color y medios de un objeto sin editar se conservan byte a byte.                                                                                                                                                                                                                                                                         |
| Tema                                                                            | Se conserva salvo que se haya cambiado la paleta del tema.                                                                                                                                                                                                                                                                                                                      |

`src/core/xlsx/write/preservation.test.ts` comprueba las primeras filas de esta tabla: construye un paquete
con un proyecto VBA, XML personalizado, un vínculo externo con un destino externo, una parte de conexiones,
configuración de impresora, una parte de propiedades de control, una relación desconocida, entradas `extLst` desconocidas
en el libro y en la hoja, cada elemento de hoja listado arriba y una segunda ventana del libro; edita
una celda, guarda y afirma que cada parte es idéntica byte a byte, que cada relación y cada tipo de contenido
está presente y que cada fragmento XML está de nuevo en el orden del esquema. Un segundo guardado del resultado es estable.
`save.test.ts` hace lo mismo con los archivos propios de Excel y de openpyxl.

## Lo que se descarta, y por qué {#what-is-dropped-and-why}

| Descartado                                                                       | Motivo                                                                                                                                                                    |
| -------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `xl/calcChain.xml`                                                               | Lo reconstruye Excel; el libro se guarda con `fullCalcOnLoad` para que cada valor se recalcule al abrirlo.                                                                |
| Firmas digitales (el origen, las partes de firma y sus relaciones)               | Cualquier cambio las invalida. Se informa al cargar: "The workbook is digitally signed ... saving removes it". Excel también las descarta.                                |
| El proyecto VBA al guardar como `.xlsx`                                          | El formato no puede contenerlo. El formato `.xlsm` se conserva tal como se cargó.                                                                                         |
| `x15ac:absPath` (dónde se guardó por última vez el archivo) y `xr:revisionPtr`   | Indicaciones de autoría para Excel, que las regenera. La ruta absoluta además revelaría la estructura de carpetas del autor.                                              |
| Atributos `xr:uid`, `spans` de fila, `x14ac:dyDescent`                           | Identificadores de revisión y pistas de diseño que Excel regenera.                                                                                                        |
| `ph` de fila y de celda (mostrar fonética), `cm` de celda sin parte de metadatos | No modelados. Las guías fonéticas se mantienen en las cadenas compartidas; solo se pierde el indicador que las muestra.                                                   |
| Fórmulas compartidas y cadenas en línea como tales                               | Se escriben como las fórmulas simples equivalentes y como cadenas compartidas (los mismos valores).                                                                       |
| Tipos de regla de formato condicional que el editor no admite                    | Se informa al cargar: `Conditional formatting rule type "..." is not supported and was dropped.` Las demás reglas de la hoja se conservan.                                |
| Una parte de tabla o de gráfico que no se puede leer                             | Se informa al cargar ("A table part could not be read and was dropped", "Chart ... could not be read; it is kept but not shown": la parte del gráfico en sí se conserva). |
| Un elemento de hoja fuera del esquema                                            | No puede colocarse en una posición válida, así que se informa al cargar: `the element "..." is not part of the worksheet schema and is dropped on save.`                  |

## Lo que se rechaza o se informa {#what-is-refused-or-reported}

Nada se rechaza en silencio. El editor informa mediante `workbook.warnings` (que los hosts muestran como aviso) y lanzando errores:

- un paquete cifrado sin contraseña pide una (`load/detect.ts`); una contraseña incorrecta es un error;
- una parte no admitida o dañada se informa por su nombre ("... is kept but not shown", "... could not be read");
- un libro con macros VBA o vínculos externos indica que se conservan pero no se ejecutan ni se actualizan;
- un libro con firma digital indica que la firma se elimina al guardar;
- un gráfico que no puede editarse sin perder detalle se marca como `editBlocked` y se muestra y guarda sin cambios;
- `.xlsb` y `.ods` no se abren; el `.xls` heredado se convierte a `.xlsx` al guardar, con un aviso por cada función que el lector binario no puede decodificar;
- editar parte de una tabla de datos se rechaza ("Cannot change part of a data table.").

## Límites conocidos de la garantía {#known-limits-of-the-guarantee}

- Una parte conservada byte a byte no se reescribe cuando se mueven las celdas a las que hace referencia. Los gráficos, las hojas de gráfico, las cachés de tablas dinámicas y las demás partes que el escritor conoce se parchean para los cambios de nombre y las eliminaciones de hojas y para los movimientos de sus rangos; una parte de XML personalizado, de conexión o externa que menciona un rango de celdas conserva el texto antiguo.
- La edición concurrente (el área de colaboración) lleva el subconjunto modelado en el documento compartido; un par sin el paquete original no puede escribir de vuelta las partes no modeladas.
- La comprobación nativa de que Excel abre un archivo guardado sin reparación es manual (`src/core/scripts/xlsx-excel-acceptance.ps1`); la fecha de la última ejecución correcta pertenece al estado del proyecto, no a esta lista.
