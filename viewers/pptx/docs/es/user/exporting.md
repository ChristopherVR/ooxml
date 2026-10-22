---
title: Exportar
description: Guarde sus ediciones de nuevo en .pptx, exporte PDF, PNG, GIF y video, empaquete para compartir e imprima documentos para el público y páginas de notas.
---

# Exportar

Todo lo que exporta se genera en su navegador y se descarga en su dispositivo; no se sube nada a ningún sitio. Las acciones de exportar y guardar están en la pestaña **Archivo** (la vista Backstage), y las más comunes se repiten en el menú de desbordamiento **"..."** en el extremo derecho de la barra de herramientas.

![File tab showing export options](/user-guide/exporting-file-menu.jpg)

## Guardar de nuevo en PowerPoint

1. Haga clic en la pestaña **Archivo**.
2. Haga clic en **Guardar** para descargar la presentación como archivo `.pptx`, o abra **Guardar como** para ver todos los formatos:

| Opción de Guardar como                          | Qué obtiene                                                                                                |
| ----------------------------------------------- | ---------------------------------------------------------------------------------------------------------- |
| **Presentación de PowerPoint**                  | Un archivo `.pptx` editable con todos sus cambios.                                                         |
| **Presentación con diapositivas de PowerPoint** | Un archivo `.ppsx` que se abre directamente en una presentación con diapositivas.                          |
| **Presentación habilitada para macros**         | Un archivo `.pptm` que conserva las macros VBA. Solo se muestra cuando el archivo abierto contiene macros. |
| **Empaquetar para compartir**                   | Un paquete `.zip` con el `.pptx` y un archivo léame.                                                       |

El botón **Guardar** de la barra de título hace lo mismo que **Archivo > Guardar**. Los archivos guardados conservan sus ediciones en el viaje de ida y vuelta: elementos añadidos y modificados, orden de las diapositivas, notas, secciones, etc.

::: warning No hay Ctrl+S
Como la aplicación se ejecuta en un navegador, **Ctrl+S** activa el cuadro de diálogo de guardado de página del propio navegador, no un guardado de la presentación. Use en su lugar el botón **Guardar** o la pestaña Archivo. Los archivos siempre se guardan como descargas; revise la carpeta de descargas de su navegador.
:::

## Exportar imágenes, PDF, GIF y video

Abra **Archivo > Exportar**:

| Opción                          | Qué obtiene                                                                  | Alcance               |
| ------------------------------- | ---------------------------------------------------------------------------- | --------------------- |
| **Crear PDF**                   | Un PDF con una página por diapositiva, capturado en alta resolución.         | Presentación completa |
| **Exportar diapositiva actual** | Una imagen PNG de la diapositiva actual.                                     | Diapositiva actual    |
| **Crear un video**              | Un video `.webm` que recorre la presentación, unos segundos por diapositiva. | Presentación completa |
| **Crear un GIF animado**        | Un GIF compacto en bucle de la presentación.                                 | Presentación completa |
| **Exportar como JSON**          | Un documento JSON portátil que se vuelve a importar con total fidelidad.     | Presentación completa |
| **Copiar como imagen**          | Copia la diapositiva actual en el portapapeles como imagen.                  | Diapositiva actual    |

Para los formatos de presentación completa, un cuadro de diálogo de progreso muestra cada diapositiva mientras se captura, con un botón **Cancelar** si cambia de opinión. El archivo terminado se descarga automáticamente.

## Imprimir

1. Abra **Archivo > Imprimir** y haga clic en **Imprimir presentación**. Se abre un cuadro de diálogo de impresión con una vista previa en vivo.
2. Elija qué imprimir en **Imprimir**: **Diapositivas de página completa**, **Documentos para el público** (con un ajuste de diapositivas por página), **Páginas de notas** (cada diapositiva sobre sus notas del orador) o **Esquema**.
3. Defina la **Orientación**, el **Modo de color** (Color, Escala de grises, Blanco y negro), si se **enmarcan las diapositivas** y el **Intervalo de diapositivas** (Todas las diapositivas, Diapositiva actual o un intervalo Desde/Hasta personalizado).
4. Haga clic en **Imprimir**. Se abre la ventana de impresión de su navegador, donde elige la impresora o guarda como PDF.

::: tip Documentos de notas
Imprimir **Páginas de notas** es la forma de obtener un documento que empareja cada diapositiva con sus notas del orador.
:::

## Proteger e inspeccionar el archivo

La página **Archivo > Información** ofrece:

- **Proteger presentación** - establezca una contraseña para que el archivo guardado esté cifrado.
- **Inspeccionar presentación** - revise las propiedades del documento.
- **Incrustar fuentes** - mantenga la tipografía coherente cuando el archivo pasa de un dispositivo a otro.
- **Firmas digitales** - vea las firmas asociadas a la presentación.

## Notas sobre fidelidad y tamaño

::: tip Las exportaciones raster coinciden fielmente con la pantalla
Las exportaciones PNG, PDF, GIF y video capturan la diapositiva en pantalla como un documento SVG autónomo y dejan que el propio motor de renderizado de su navegador lo dibuje, de modo que efectos como fondos desenfocados/transparentes, colores personalizados y biseles 3D se conservan en lugar de aproximarse. `html2canvas` solo se usa como respaldo, para el raro recurso (una imagen de otro origen sin credenciales, o una fuente que no se carga) que no se puede capturar de este modo. Consulte [Limitaciones](/es/guide/limitations) para el pequeño conjunto de aproximaciones CSS restantes.
:::

::: info Diapositivas muy grandes
Una exportación cuya resolución superaría el tamaño máximo de lienzo de su navegador se divide automáticamente en mosaicos y se reensambla en el archivo final, de modo que una diapositiva extremadamente grande exportada a una escala alta sigue produciendo un PNG, PDF, GIF o video completo y del tamaño correcto en lugar de recortarse o cortarse. No hay ningún ajuste que cambiar; esto ocurre de forma transparente.
:::

## A continuación

- Coedite con otras personas: [Colaboración](/es/user/collaboration)
- Detalles de exportación para desarrolladores: [/react/export](/react/export)
