# Primeros pasos

Abra un diagrama en local, conozca los límites actuales e incruste el mismo visor en su aplicación.

::: warning Compatibilidad inicial del visor, no certificación de Visio
Esta es una beta pública. No tome la apertura correcta de un archivo como prueba de que todas las formas, fórmulas o funciones del formato se representan con fidelidad. Consulte el [registro de capacidades](/parity).
:::

## Pruebe el área de pruebas local {#try-the-local-playground}

1. [Abra el área de pruebas](/demo/){target="_self"} para ver un documento de ejemplo original.
2. Elija un archivo `.vsdx` local con el selector de archivos o la zona de arrastre.
3. Cambie de página, ajuste el zoom, encaje la página y seleccione una forma.
4. Revise las notas de compatibilidad. Compare los archivos importantes con su representación en Microsoft Visio.

No se sube ningún documento. Los controles admiten búsqueda, capas de visualización, edición experimental de texto sin formato, deshacer y rehacer acotados, exportación a SVG y descargas explícitas de copias VSDX. Las instantáneas de impresión son solo una preparación de la API. Archivo > Exportar puede crear un PDF de páginas como imagen, pero el PDF vectorial o con texto seleccionable, el dibujo general y la mayoría de las funciones del formato binario heredado `.vsd` quedan fuera de esta compilación.

Las compilaciones del mismo área de pruebas para los demás frameworks figuran en [la página de demostraciones](/demos).

## Instale un paquete de framework {#install-a-framework-package}

Instale un único paquete de visor para su framework. Cada uno incluye el visor compartido, su adaptador, los workers y la API de documentos.

::: code-group

```bash [React]
npm install visio-react-viewer react
```

```bash [Vue]
npm install visio-vue-viewer vue
```

```bash [Angular]
npm install visio-angular-viewer
```

```bash [Svelte]
npm install visio-svelte-viewer svelte
```

```bash [Solid]
npm install visio-solid-viewer solid-js
```

```bash [Vanilla JS]
npm install visio-vanilla-viewer
```

:::

| Framework | Paquete                | Exportación               |
| --------- | ---------------------- | ------------------------- |
| Vanilla   | `visio-vanilla-viewer` | `mountViewer`             |
| React     | `visio-react-viewer`   | `VisioViewer`             |
| Vue 3     | `visio-vue-viewer`     | `VisioViewer`             |
| Angular   | `visio-angular-viewer` | `VisioViewerComponent`    |
| Svelte 5  | `visio-svelte-viewer`  | `VisioViewer` por defecto |
| Solid     | `visio-solid-viewer`   | `VisioViewer`             |

Los seis adaptadores reenvían el mismo controlador, las mismas propiedades, los mismos eventos y el mismo manejador imperativo a través del visor compartido. Consulte las [guías de frameworks](/frameworks/react) y el [contrato de los enlaces](/bindings).

## Monte un visor {#mount-a-viewer}

El contenedor debe tener una altura. A `load` se le puede pasar un `File`, `Blob`, `Uint8Array` o `ArrayBuffer` del navegador.

```ts
import { mountViewer } from 'visio-vanilla-viewer';

const viewer = mountViewer(host, {
	events: { 'shape-select': (shape) => console.log(shape) },
});

await viewer.load(file); // A local File or bytes
viewer.fit();

// When your view unmounts:
viewer.destroy();
```

## Ejecute el espacio de trabajo de desarrollo {#run-the-development-workspace}

El visor forma parte del monorrepositorio `ooxml`, un espacio de trabajo de Bun. Instale una vez en la raíz del repositorio con `bun install`; `ooxml-core` (`src/core/visio`) y `ooxml-ui` (`src/ui/src/visio`) se resuelven a los paquetes de la misma copia, por lo que no hace falta ningún paso de enlace.

```bash
bun install
cd viewers/visio
bun run check
bun run dev
```

Use la URL local que muestra Vite. Esto inicia un servidor de desarrollo; no publica el sitio.

## Edición experimental de texto sin formato {#editing}

Cargue un paquete de origen local, seleccione una forma compatible y use sus controles de texto. `replacePlainText(pageId, shapeId, text)`, `undo()` y `redo()` exponen las mismas operaciones a través de cada manejador. Un modelo asignado con `document` no tiene bytes de paquete y permanece en solo lectura.

La edición rechaza el texto vinculado a una plantilla maestra, el texto enriquecido, los campos, los paquetes firmados y el contenido con macros. El historial está acotado y puede descartar estados de deshacer antiguos. Las cachés de fórmulas guardadas no se recalculan. No se ha implementado la edición general de dibujo ni de estilos.

`exportVsdx()` devuelve los bytes del paquete original o editado y los diagnósticos para una descarga explícita de una copia. Nunca sobrescribe el archivo de entrada. La reapertura nativa en Visio y la fidelidad de un viaje de ida y vuelta sin pérdidas siguen sin verificarse.

## Exporte un SVG estático {#export}

`viewer.exportSvg()` devuelve la cadena SVG de la página actual, sus dimensiones, su tamaño en bytes y los diagnósticos de compatibilidad. No modifica la selección, el zoom ni el documento de origen. El botón Exportar SVG del área de pruebas realiza una descarga local explícita.

Se incrustan el contenido de fondo compatible y las imágenes rasterizadas validadas. Las fuentes no se incrustan; el ajuste de texto y otras aproximaciones siguen siendo visibles en la descripción y los metadatos del SVG. La exportación está limitada a 16 MiB, con comprobaciones tempranas y prudentes de recursos. Quien llama puede reducir el límite con `{ maxBytes }`. Esto no es una salida VSDX editable. Los rasterizadores posteriores deben limitar por su cuenta las dimensiones y el área en píxeles.

## Prepare instantáneas de página {#prepare-page-snapshots}

`createPrintSnapshot()` devuelve, a través de la API, artefactos SVG de página inmutables y diagnósticos. Conserva las dimensiones del dibujo, pero no aplica la configuración de impresora nativa, no abre un cuadro de diálogo de impresión ni implementa la exportación a PDF. Prepara los datos para un futuro flujo de impresión.

## Propiedades y eventos {#properties-and-events}

| Superficie    | Contrato actual                                                                                                                                                        |
| ------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Propiedades   | `document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `ribbonAddIns`, `collaboration`                                                                             |
| Eventos       | `document-load`, `document-change`, `document-error`, `page-change`, `zoom-change`, `shape-select`, `selection-change`, `office-ribbon-add-in`, `collaboration-change` |
| Ciclo de vida | `update`, `load`, `fit`, `exportSvg`, `createPrintSnapshot`, `replacePlainText`, `undo`, `redo`, `cancelEdit`, `exportVsdx`, `destroy`                                 |

`canEdit` es el permiso de edición (el visor es editable por defecto; `false` lo deja en solo lectura), `collaboration` comparte el dibujo abierto en una sala y `document-change` informa `kind` como `edit`, `undo`, `redo`, `remote` o `save`. Cuando el anfitrión ha guardado los bytes de `exportVsdx()`, `controller.markClean()` borra el indicador de cambios. Para componer su propia interfaz con las piezas del visor, consulte los [bloques de construcción](/building-blocks); para compartir un dibujo, consulte [Colaboración](/es/collaboration).

Monte el visor cuando el anfitrión ya exista y llame a `destroy()` cuando la vista del framework se desmonte. Destruirlo varias veces es seguro. Las actualizaciones parciales dejan sin cambios las propiedades omitidas. La referencia completa es la [API del visor](/es/api).

## Verificación antes de la adopción {#verification-before-adoption}

Ejecute `bun run check` y `bun run test:browser` en `viewers/visio`. Las pruebas del analizador del núcleo están junto al motor en `src/core/visio` (se ejecutan con `bun run test` en la raíz del repositorio). Estas comprobaciones establecen únicamente el comportamiento que afirman; no establecen la paridad visual con Microsoft Visio.
