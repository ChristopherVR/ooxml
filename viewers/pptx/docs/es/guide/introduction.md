---
title: ¿Qué es pptx-viewer?
description: Una descripción general del monorepo TypeScript pptx-viewer para analizar, editar, renderizar y convertir archivos PowerPoint con React, Vue 3, Angular, Svelte 5 y JavaScript vanilla.
---

# ¿Qué es pptx-viewer?

`pptx-viewer` es un monorepo TypeScript para **analizar, editar, renderizar y convertir** archivos Microsoft PowerPoint (`.pptx`), en el navegador y en Node.js. Funciona completamente en memoria sobre el archivo ZIP OpenXML, sin dependencias nativas.

`pptx-viewer` cubre el ciclo completo: cargar una presentación existente, mutar su modelo de datos estructurado, renderizarla y guardarla de vuelta en un archivo `.pptx` válido. El mismo motor core impulsa componentes para **React**, **Vue 3**, **Angular** y **Svelte 5**, además de una compilación **JavaScript vanilla** para proyectos sin framework.

## Qué hace

El SDK proporciona nueve capacidades principales:

1. **Analizar** archivos `.pptx` (y `.ppsx` / `.pptm` / `.potx`, además de `.ppt` binario heredado de PowerPoint 97-2003) desde un `ArrayBuffer` sin procesar a un modelo [`PptxData`](/es/guide/data-model) estructurado.
2. **Crear** presentaciones desde cero con una API de construcción fluida.
3. **Renderizar** diapositivas como componentes interactivos de React, Vue, Angular, Svelte o vanilla JS.
4. **Editar** presentaciones de forma programática o mediante el editor WYSIWYG integrado.
5. **Guardar** los cambios de vuelta en un archivo `.pptx` válido (seguro para el viaje de ida y vuelta).
6. **Convertir** presentaciones a Markdown con extracción opcional de medios.
7. **Exportar** diapositivas como imágenes (PNG/JPEG), SVG, PDF, GIF o video.
8. **Colaborar** en tiempo real mediante Yjs CRDT con seguimiento de presencia.
9. **Cifrar/Descifrar** archivos PPTX protegidos con contraseña (AES-128/256).

El motor maneja elementos de presentación, cientos de formas predefinidas `ST_ShapeType`, gráficos, estilos de tabla integrados, SmartArt, modelos 3D, animaciones y transiciones (incluido morph), temas, patrones de diapositivas, medios incrustados, metarchivos EMF/WMF, objetos OLE, tinta digital, firmas digitales, cifrado, preservación de macros VBA y conformidad OOXML Strict.

## Para quién está pensado

- **Equipos de producto que integran un visor de documentos**: muestre presentaciones subidas por los usuarios dentro de una aplicación web sin enviar archivos a un servicio de conversión.
- **Equipos que crean un editor**: los enlaces incluyen una superficie de edición WYSIWYG completa (cinta, inspector, deshacer/rehacer, colaboración) que puede activar por funcionalidad.
- **Trabajo de backend y automatización**: el paquete core es independiente del framework y se ejecuta en Node.js, por lo que puede generar, inspeccionar, comparar o convertir presentaciones en scripts, servidores y CI.
- **Canalizaciones de IA y agentes**: `pptx-viewer-mcp` expone el motor como llamadas a herramientas MCP, y el conversor Markdown convierte las presentaciones en texto apto para LLM.

## Modelo de renderizado

`pptx-viewer` no rasteriza las diapositivas, ni en el servidor ni sobre un `<canvas>`. Las diapositivas se renderizan como **HTML, CSS y SVG en el DOM**:

- El texto es texto HTML real: seleccionable, buscable, traducible y legible por lectores de pantalla.
- Las formas y los conectores son SVG con trazados de recorte calculados por el motor de geometría.
- Las tablas son elementos HTML `<table>`; los gráficos son SVG en línea construido a partir de los datos del gráfico analizado.
- El zoom es una transformación CSS, por lo que las diapositivas se mantienen nítidas a cualquier escala y densidad de pantalla.
- Como cada elemento es un nodo DOM, el editor obtiene selección, tiradores de arrastre, edición de texto en línea y accesibilidad sin esfuerzo.

La rasterización solo ocurre cuando usted la pide: la exportación raster (PNG/JPEG/PDF/GIF/video) intenta primero una captura SVG `foreignObject` (el propio motor del navegador pinta el DOM, conservando `backdrop-filter`, propiedades personalizadas y transformaciones 3D), recurre a volver a rasterizar la canalización de exportación SVG vectorial existente, y solo como último recurso dibuja el DOM mediante `html2canvas-pro`.

::: tip Compromisos
CSS no puede expresar todos los efectos de PowerPoint con precisión de píxel, y la ruta de respaldo de `html2canvas-pro` (usada cuando un recurso no puede incrustarse para la captura SVG) tiene sus propias carencias, por ejemplo aproximar los modos de fusión con opacidad. Consulte [Limitaciones](/es/guide/limitations) para ver la lista completa.
:::

## Los paquetes

El monorepo publica ocho paquetes.

| Paquete                                 | Nombre npm                   | Propósito                                                                                                          |
| --------------------------------------- | ---------------------------- | ------------------------------------------------------------------------------------------------------------------ |
| [**Core**](/core/)                      | `pptx-viewer-core`           | Analizar, crear, editar, serializar y convertir archivos PPTX. Independiente del framework.                        |
| [**React**](/react/)                    | `pptx-react-viewer`          | Visualizador, editor y presentador React con barra de herramientas, inspector, colaboración y exportación.         |
| [**Vue 3**](/vue/)                      | `pptx-vue-viewer`            | Visualizador/editor Vue 3 construido sobre el mismo motor, con el mismo conjunto de funciones.                     |
| [**Angular**](/angular/)                | `pptx-angular-viewer`        | Visualizador/editor Angular construido sobre el mismo motor, con el mismo conjunto de funciones.                   |
| [**Vanilla JS**](/vanilla/)             | `pptx-vanilla-viewer`        | Enlace sin framework construido sobre el mismo motor: DOM puro, una función fábrica, sin dependencia de framework. |
| [**Svelte**](/svelte/)                  | `pptx-svelte-viewer`         | Componente Svelte 5 construido sobre el mismo motor, con el mismo conjunto de funciones.                           |
| [**Herramientas / MCP**](/packages/mcp) | `pptx-viewer-mcp`            | 73 funciones de herramientas PPTX, un servidor MCP para agentes de IA y el códec de colaboración Y.Doc.            |
| **Instalador**                          | `@christophervr/pptx-viewer` | CLI interactivo que integra el paquete de visualizador correcto en su proyecto.                                    |

### Cómo encajan las piezas

Los cinco paquetes de enlace de interfaz se apoyan en una capa de renderizado compartida, que a su vez se apoya en Core:

```
pptx-react-viewer   ┐
pptx-vue-viewer     │
pptx-angular-viewer ├── ooxml-ui/pptx ── ooxml-core/pptx
pptx-vanilla-viewer │                               ├── emf-converter
pptx-svelte-viewer  ┘                               └── mtx-decompressor
```

- **`pptx-viewer-core`** posee todo lo relativo al formato de archivo: las canalizaciones de carga y guardado, el modelo de datos tipado, la resolución del tema, el motor de geometría, el cifrado y los conversores Markdown/SVG. No tiene interfaz y se ejecuta en cualquier lugar donde se ejecute JavaScript.
- **`ooxml-ui/pptx`** publica el renderizado compartido, los estilos, los web components, la carga en el navegador y las vistas de exportación desde `src/ui/src/pptx`. Todos los enlaces importan esta API pública. Las operaciones de edición y colaboración sin DOM viven en `ooxml-core/pptx/editor`; el paquete privado `pptx-viewer-shared` es solo una fachada de compatibilidad.
- **Los enlaces** son capas de vista delgadas. Cada uno traduce los mismos datos de renderizado compartidos al idioma de su framework (JSX, plantillas SFC, plantillas de Angular, runes de Svelte o DOM simple), por lo que los cinco enlaces renderizan de forma idéntica y se mantienen en paridad de funciones.
- **`emf-converter`** (metarchivo EMF/WMF a PNG) y **`mtx-decompressor`** (fuentes incrustadas MicroType Express) son dependencias npm independientes de Core.
- **`pptx-viewer-mcp`** se apoya en `pptx-viewer-core` para exponer superficies de llamadas a herramientas y de colaboración a agentes de IA.

## Ejemplos mínimos

::: code-group

```ts [Load, edit, save]
import { PptxHandler } from 'pptx-viewer-core';

const handler = new PptxHandler();
const data = await handler.load(arrayBuffer);

// data.slides is a typed, mutable model
data.slides[0].elements.filter((el) => el.type === 'text').forEach((el) => console.log(el.text));

const bytes = await handler.save(data.slides); // => Uint8Array (.pptx)
```

```ts [Create from scratch]
import { PptxHandler } from 'pptx-viewer-core';

const { handler, data, createSlide } = await PptxHandler.create({
	title: 'Q4 Report',
	theme: { colors: { accent1: '#FF6B6B' } },
});

data.slides.push(createSlide('Blank').addText('Hello', { fontSize: 36 }).build());

const bytes = await handler.save(data.slides);
```

:::

## Siguientes pasos

- [Instalación](/es/guide/installation) - instalar los paquetes y configurar el desarrollo local.
- [Inicio rápido](/es/guide/quick-start) - ejemplos de uso de extremo a extremo.
- [Arquitectura](/es/guide/architecture) - cómo funcionan las canalizaciones de carga/guardado y el diseño por capas.
- [El modelo PptxData](/es/guide/data-model) - el modelo tipado que usted edita.
- [Resumen del paquete Core](/core/) - el motor de análisis, edición y serialización.
- [Resumen del paquete React](/react/) - el componente visualizador/editor.
- [Limitaciones](/es/guide/limitations) - advertencias importantes a leer antes de adoptarlo.
