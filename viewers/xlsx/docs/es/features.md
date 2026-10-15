# Funciones y limitaciones

Esta es una implementación temprana. **No tiene paridad con Microsoft Excel** y el guardado no es sin pérdidas para las funciones que el modelo no cubre. Esta página enumera lo que el editor está preparado para gestionar y lo que no; cuando un libro contiene algo no admitido, el editor lo notifica (`workbook-warning`) en lugar de ocultarlo.

El [estado actual de la paridad y las pruebas](https://github.com/ChristopherVR/ooxml/blob/main/docs/xlsx-parity-status.md)
resumen el trabajo implementado, las comparaciones con Excel nativo y las carencias pendientes.

::: warning Versión temprana
El motor de libros de `ooxml-core` y el componente `<xlsx-editor>` son jóvenes. Las pruebas de navegador (`e2e/xlsx/*.spec.ts`) cubren la apertura de `.xlsx`, `.xls` y `.csv`, la escritura de valores y fórmulas, el recálculo, el formato desde la cinta, la inserción y eliminación de filas, las pestañas de hoja, deshacer y rehacer, Buscar y reemplazar, Formato de celdas, la selección de gráficos, guardar y reabrir, el modo de solo lectura, las configuraciones regionales y los seis enlaces de framework. Todo lo demás que se indica a continuación está implementado pero menos probado; cabe esperar asperezas.
:::

## Formatos de archivo

| Formato                  | Abrir | Guardar      | Notas                                                                                                                                                                                                                                                                          |
| ------------------------ | ----- | ------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `.xlsx`, `.xltx`         | sí    | sí           | El formato principal.                                                                                                                                                                                                                                                          |
| `.xlsm`                  | sí    | sí           | El proyecto VBA se transmite sin cambios; las macros nunca se ejecutan.                                                                                                                                                                                                        |
| `.xls` (97-2003)         | sí    | como `.xlsx` | Se lee mediante los códecs compartidos de ole2: valores, fórmulas cuando se pueden decodificar, formato básico, celdas combinadas, anchuras de columna, alturas de fila, paneles inmovilizados.                                                                                |
| `.csv`                   | sí    | sí           | RFC 4180, delimitador detectado automáticamente; los campos que empiezan por = siguen siendo texto al importar. La exportación escribe los valores mostrados de la hoja activa y protege contra la inyección de fórmulas.                                                      |
| Protegido con contraseña | sí    | sí           | Los archivos .xlsx/.xlsm cifrados solicitan una contraseña. Archivo > Información > Cifrar libro establece una contraseña para los guardados posteriores de Excel. Al abrir un archivo cifrado no se conserva su contraseña para guardar; vuelva a definirla antes de guardar. |
| `.xlsb`, `.ods`          | no    | no           |                                                                                                                                                                                                                                                                                |

## Cuadrícula y formato

- Valores de celda, cadenas compartidas con tramos de texto enriquecido, booleanos, errores y fechas (sistemas de fechas de 1900 y 1904).
- Fuentes, rellenos (sólido, de trama, degradado aproximado), bordes, alineación, ajuste de texto, sangría, rotación.
- Formatos numéricos: secciones, condiciones, colores, fechas y horas, fracciones, científico, texto. Un número General que no cabe en su columna pierde decimales o pasa a notación científica, y los demás números muestran `###`, medido con las fuentes del navegador.
- Las alturas de fila siguen al texto ajustado o ampliado tras una edición, salvo que la fila tenga una altura personalizada.
- Celdas combinadas, anchuras de columna y alturas de fila, filas y columnas ocultas, paneles inmovilizados, zoom, líneas de cuadrícula.
- Formato condicional: reglas de celda, reglas de texto, "Una fecha que ocurra" (hoy, últimos 7 días, este mes, ...), superiores e inferiores, por encima del promedio, duplicados, escalas de color, barras de datos, conjuntos de iconos, con un Administrador de reglas que edita, reordena y elimina reglas.
- Validación de datos, con la lista desplegable ofrecida en la cuadrícula.
- Comentarios (se muestran y se editan; las respuestas en hilo se muestran como notas), hipervínculos.

## Fórmulas

El motor de cálculo vive en `ooxml-core` y recalcula las celdas dependientes tras cada edición. Su catálogo contiene 481 funciones (contadas a partir de `FUNCTION_CATALOG` el 2026-10-08) de matemáticas y trigonometría, estadística (incluidas las distribuciones inversas y la familia `LINEST`, `LOGEST`, `TREND` y `GROWTH`), lógica, texto (incluidas `REGEXTEST`, `REGEXEXTRACT` y `REGEXREPLACE`), fechas, búsqueda (incluidas `XLOOKUP` y matrices dinámicas como `FILTER`, `SORT`, `UNIQUE` y `SEQUENCE`), información, financieras (incluidas las funciones de bonos, cupones, valores con descuento y letras del Tesoro, `VDB` y `XIRR`), ingeniería (incluidas `COMPLEX`, las funciones `IM` y `CONVERT`) y bases de datos (`DSUM`, `DCOUNT`, `DGET` y el resto). Insertar función las enumera por categoría.

No existen todas las funciones de Excel: una desconocida muestra `#NAME?`. No están implementadas: `ODDFPRICE`, `ODDFYIELD`, `ODDLPRICE`, `ODDLYIELD`, `AMORDEGRC`, `BESSELI`, `BESSELJ`, `BESSELK`, `BESSELY`, `FORECAST.ETS`, `ASC`, `DBCS`, `JIS`, `PHONETIC`, `BAHTTEXT`, `INFO`, `GETPIVOTDATA`, `ISOMITTED` y `WEBSERVICE`. Las referencias a libros externos se conservan pero no se resuelven.

Las referencias circulares se notifican y se dejan en 0 a menos que el cálculo iterativo esté activado. Actívelo en Archivo > Opciones > Fórmulas (Habilitar cálculo iterativo, con Iteraciones máximas y Cambio máximo); el ajuste se guarda con el libro (`calcPr iterate`, `iterateCount`, `iterateDelta`). Un grupo circular se recalcula entonces hasta el máximo de iteraciones y se detiene cuando ninguna celda cambia en más del cambio máximo. Activarlo o desactivarlo no es un paso de deshacer.

El cálculo se puede alternar entre Automático y Manual (Fórmulas > Opciones para el cálculo, o Archivo > Opciones); el modo se guarda con el libro. Calcular ahora (F9) y Calcular hoja (Mayús+F9) recalculan a petición. Mostrar fórmulas (Ctrl+\`) muestra el texto de las fórmulas en lugar de los resultados.

## Edición

Escritura en las celdas (valores y fórmulas), la barra de fórmulas y el cuadro de nombres, Ctrl+; y Ctrl+Mayús+; para la fecha y la hora actuales, rellenar, copiar y pegar (con el portapapeles del sistema como texto y HTML), insertar y eliminar filas, columnas y celdas, ordenar, filtrar, quitar duplicados, buscar y reemplazar, gestión de hojas (añadir, cambiar el nombre, mover, ocultar, color; las hojas nuevas reciben nombre en el idioma de la interfaz, por ejemplo `Tabelle2` en alemán), los estilos de celda integrados, tablas (opciones de estilo, filas de encabezado y de totales, cambio de tamaño, convertir en rango), grupos de esquema de filas y columnas, configuración de página y opciones de impresión (líneas de cuadrícula, encabezados, centrado), deshacer y rehacer con nombres de paso significativos. Las referencias de las fórmulas siguen a las ediciones estructurales.

La protección de hojas y de la estructura del libro puede llevar una contraseña. Impide ediciones accidentales; no es seguridad. El editor lee y comprueba tanto el hash heredado de 16 bits de Excel como su hash más reciente SHA-512 (también SHA-1, SHA-256 y SHA-384), de modo que Desproteger hoja y Desproteger libro piden la contraseña sea cual sea el hash que lleve el archivo. Una contraseña definida en el editor se escribe como el hash heredado. Una protección cuyo hash use un algoritmo que el núcleo no puede calcular nunca se desbloquea con una contraseña.

Las imágenes y los gráficos se seleccionan haciendo clic en ellos: se mueven y cambian de tamaño con el ratón, se desplazan con las teclas de flecha y se eliminan con Supr; un gráfico seleccionado muestra la pestaña Diseño de gráfico.

## Gráficos y SmartArt

Las familias de gráficos más comunes (columnas, barras, líneas, áreas, circular, anillo, dispersión y radial)
se dibujan como SVG a partir de valores en vivo. La inserción de gráficos, el tipo y la agrupación, el título, la leyenda y los
cambios de color usan el núcleo compartido y los controles de galería de Office. Los paneles de formato acoplados
ofrecen rellenos de series, el espaciado de barras admitido y rellenos para el área del gráfico,
el área de trazado, el título y la leyenda. Los tipos de degradado admitidos, las paradas, la transparencia,
el brillo, las vistas previas de dirección y los preajustes nativos comparten los mismos controles de relleno.

Los estilos de gráfico importados, las sombras y degradados admitidos, la visibilidad de los ejes, las fuentes
heredadas, los tramos de título con formato mixto, el ajuste automático del título y el espaciado entre párrafos del título de punto o porcentaje
cuentan con compatibilidad de representación y conservación. Las posiciones manuales de título importadas y
las superposiciones de título se representan; el XML de diseño de título, área de trazado y leyenda se conserva
al cambiar el tipo de gráfico. Los rectángulos manuales de leyenda importados y las superposiciones se representan,
incluidas las disposiciones horizontales y verticales de entradas probadas. Las dimensiones manuales del título,
la ubicación completa del área de trazado, la geometría automática de la leyenda, el empaquetado de leyendas en varias filas o columnas y
los controles de creación de posiciones siguen incompletos.
Los gráficos de burbujas, de cotizaciones y de superficie no se dibujan: se representan como un marco
con etiqueta y se conservan al guardar (`src/core/xlsx/layout/chart-svg.ts`,
`chart.test.ts`). La creación de estilos de gráfico no está admitida: las definiciones
de estilo importadas se conservan y se representan, pero no se pueden crear ni editar
(`src/core/xlsx/model.ts`, `styleDefinition`). La representación y el formato
anteriores están cubiertos por las pruebas de gráficos del núcleo y por `docs/xlsx-parity-review.md`
en el repositorio ooxml.
Las comparaciones con Excel nativo verifican rutas concretas de edición y exportación. La medición exacta del texto del gráfico,
el ajuste, la ubicación y la fidelidad de trama siguen incompletos, al igual que
muchos controles de ejes, etiquetas y efectos, las familias de gráficos avanzadas y la interfaz nativa completa.

SmartArt usa el renderizador compartido de dibujos en caché con las fuentes y los
colores del tema del libro. Sigue siendo solo de visualización: no están implementados la inserción, la edición de texto, el reflujo
ni las galerías completas de diseño, color y estilo. Si falta un dibujo en caché
se produce un marcador de posición de texto.

## Minigráficos

Los minigráficos de Excel 2010 se dibujan en sus celdas anfitrionas, detrás del texto de la celda:
línea (con marcadores y puntos máximo, mínimo, primero, último y negativos), columna y
ganancia/pérdida. El dibujo sigue los ajustes del grupo para las celdas vacías (vacío, cero o
conectado), el eje vertical (cada minigráfico, todo el grupo o límites personalizados),
el eje horizontal en cero, el eje de fechas, el trazado de derecha a izquierda, las celdas ocultas
y los colores de tema o RGB. Un minigráfico se vuelve a dibujar cuando cambia una celda de su rango de datos
y cambia de tamaño con su celda. El XML del minigráfico se guarda exactamente como se cargó,
y las filas y columnas insertadas o eliminadas lo mueven con sus celdas.

Insertar > Minigráficos crea grupos de línea, columna y ganancia/pérdida a partir de un rango de datos
y un rango de ubicación. Un rango de ubicación vertical asigna las filas de datos a minigráficos
individuales; un rango de ubicación horizontal asigna las columnas de datos. Seleccione una celda anfitriona
y elija Configuración de minigráficos para cambiar el tipo, los marcadores, los colores de serie y de valores negativos,
el grosor de línea, el tratamiento de celdas vacías y ocultas, la visibilidad de los ejes, la dirección
y los límites verticales individuales, compartidos o personalizados de todo el grupo.
Borrar minigráficos seleccionados quita los anfitriones seleccionados; Borrar grupos de minigráficos
quita sus grupos completos. Estos cambios admiten deshacer y guardar/cargar.

Las galerías de estilos, el agrupamiento y desagrupamiento y la creación de ejes de fechas no están disponibles.
Los ejes de fechas existentes se conservan. No se reproduce el resultado exacto en píxeles de Excel (tamaño de marcador, separación de columnas y
suavizado de líneas), y las celdas no numéricas de un rango de datos se
tratan como vacías.

## No admitido (se conserva siempre que es posible)

| Función                                    | Comportamiento                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      |
| ------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Tablas dinámicas y cachés dinámicas        | Se conservan al guardar; no se muestran como tablas dinámicas y no se actualizan.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   |
| Segmentaciones de datos, escalas de tiempo | Se conservan al guardar; no se dibujan.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             |
| Creación de minigráficos                   | La creación, los ajustes de grupo y el borrado están admitidos (véase arriba); las galerías de estilos, el agrupamiento y desagrupamiento y la creación de ejes de fechas no están implementados.                                                                                                                                                                                                                                                                                                                                                                   |
| Gráficos avanzados y opciones de gráfico   | Las familias comunes y el formato admitido se describen arriba. Las familias avanzadas, la creación completa de ejes, etiquetas y efectos y el diseño nativo exacto siguen incompletos.                                                                                                                                                                                                                                                                                                                                                                             |
| Formas, SmartArt, controles de formulario  | Las imágenes se muestran. SmartArt usa el dibujo en caché mediante el renderizador compartido (solo visualización); sin dibujo en caché se enumera su texto. Los demás dibujos se muestran como marcadores de posición.                                                                                                                                                                                                                                                                                                                                             |
| Macros (VBA), complementos, Power Query    | Nunca se ejecutan; el proyecto VBA se transmite.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| Vínculos externos y conexiones de datos    | Se conservan; los valores son los almacenados en caché.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             |
| Colaboración en tiempo real                | Los valores de celda, las fórmulas, los formatos, las hojas, los tamaños de fila y columna, las celdas combinadas y los nombres se comparten mediante `ooxml-core/collab` (Archivo > Compartir o la propiedad `collaboration`); las selecciones remotas se muestran en la cuadrícula y deshacer es por persona. Los gráficos, los dibujos, los comentarios, los formatos condicionales, la validación, las tablas y la configuración de página permanecen locales; el anfitrión se encarga del servidor y del guardado. Consulte [colaboración](/es/collaboration). |
| Impresión                                  | La impresión del navegador del rango usado o del área de impresión; sin vista Diseño de página ni vista previa de salto de página.                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| Propiedades del documento                  | Las propiedades principales, de empresa, de administrador y personalizadas se pueden editar en Archivo > Información con deshacer. Los tipos de propiedad personalizada desconocidos se conservan y se muestran en solo lectura.                                                                                                                                                                                                                                                                                                                                    |

Las pruebas de ida y vuelta del núcleo y las comprobaciones de aceptación con Excel están en el repositorio `ooxml`. Notifique un libro que se represente o guarde mal como una incidencia, adjuntando el archivo si puede compartirlo.

Las firmas digitales se notifican al abrir un libro y se eliminan al guardar. El editor no firma archivos. Los hipervínculos siguen la política del núcleo compartido; solo los enlaces web y de correo electrónico se abren externamente.
