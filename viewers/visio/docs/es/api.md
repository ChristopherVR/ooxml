# API del visor

La API independiente del framework que hay detrás de cada paquete de visor. Cada paquete de framework
(`visio-react-viewer`, `visio-vue-viewer`, `visio-angular-viewer`,
`visio-svelte-viewer`, `visio-solid-viewer`, `visio-vanilla-viewer`) reenvía
los mismos métodos y eventos del manejador; consulte `packages/bindings/README.md` para ver las
propiedades, eventos y ciclo de vida nativos de cada adaptador. El renderizado y la edición son
funciones en beta; el [registro de capacidades](/parity) indica lo que falta.

## Montar y cargar {#mount-and-load}

```ts
import { mountViewer } from 'visio-vanilla-viewer';
const viewer = mountViewer(container, {
	zoom: 1,
	events: { 'document-error': (error) => console.error(error) },
});
await viewer.load(file); // Blob, Uint8Array or ArrayBuffer
viewer.fit();
viewer.update({ pageIndex: 1 });
const snapshot = viewer.exportSvg(); // SVG string, dimensions, byteLength and diagnostics
const prepared = viewer.createPrintSnapshot(); // Frozen current-page artifacts, no printing
viewer.destroy(); // idempotent; later commands reject
```

Importar el módulo es seguro para SSR; el montaje y el registro del elemento
solo ocurren en el navegador. `ViewerController` no depende del DOM, de modo que otro renderizador puede suscribirse
al mismo estado y a los mismos eventos. `renderPage` es el puerto del renderizador SVG. Consulte la
[arquitectura](/es/architecture) para ver la responsabilidad de cada módulo.

## Propiedades y eventos {#properties-and-events}

Todos los bindings aceptan las mismas propiedades y notifican los mismos eventos (el contrato son `ViewerProperties` y `ViewerEvents` en `ooxml-ui/visio`; [bindings](/bindings) lista las formas nativas).

| Superficie  | Contrato                                                                                                                                                                               |
| ----------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Propiedades | `document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `readOnly`, `ribbonAddIns`, `collaboration`, `storage`, `customization`                                                     |
| Eventos     | `document-load`, `document-change`, `document-error`, `page-change`, `zoom-change`, `shape-select`, `selection-change`, `office-ribbon-add-in`, `collaboration-change`, `file-command` |

- `canEdit` (booleano) es el permiso de edición. El visor completo es editable por defecto y `canEdit: false` lo deja en solo lectura; los [bloques de construcción](/building-blocks) son de solo lectura hasta que el anfitrión pase `true`. Un dibujo sin bytes de paquete (un modelo pasado como `document`, un `.vsd` heredado) es de solo lectura sea cual sea `canEdit`.
- `document-change` entrega `{ document, dirty, kind }`. `kind` es `edit`, `undo` o `redo` para los cambios de esta ventana, `remote` cuando se aplicó el cambio de un colaborador y `save` tras `markClean()`.
- `markClean()` (en `viewer.controller` y en los bloques de construcción) declara guardado el contenido actual. Un anfitrión que guarda `exportVsdx()` por su cuenta lo llama después: `edit.dirty` pasa a `false`, se emite un `document-change` de tipo `save`, y las ediciones y pasos de deshacer posteriores se miden desde ese punto. No hace nada con un dibujo que no se puede guardar de vuelta.
- `collaboration` (`VisioCollaborationOptions | null`) comparte el dibujo abierto en una sala mientras esté definida. `collaboration-change` entrega un `CollaborationState` cuando el uso compartido empieza o termina y cuando cambian el estado, la sincronización o las personas. En el elemento (`viewer.element` de un handle), `collaborationState` lee el estado actual, y `reconnectCollaboration()` y `stopCollaboration()` actúan sobre él. Véase [Colaboración](/es/collaboration).

## Seleccionar formas {#select-shapes}

Todo manejador montado o nativo expone `selectShapes(targets)`, `selectAll()` y
`clearSelection()`. Un destino tiene la forma `{ id, name, pageId? }`; si se omite `pageId`,
se entiende la página actual. `controller.state.selectedShapes` es una matriz ordenada e
inmutable, y su primer elemento es la `selectedShape` principal. El evento
`shape-select` ya existente informa de ese elemento principal; `selection-change` informa de la
matriz completa. Estos estados están disponibles a través del hook, composable, señal o almacén nativo
de cada framework. Vue emite `selection-change`; Angular expone
`selectionChange`; los seis adaptadores aceptan `events['selection-change']`.

Hacer clic con Mayús, Ctrl o Meta alterna una forma. Los clics en el fondo y la tecla Escape
borran la selección. Los destinos ocultos o eliminados se descartan, seleccionar un grupo excluye a
sus descendientes, y cambiar de página o de documento borra la selección. Seleccionar todo incluye las
formas visibles de primer plano. Los paneles de texto y geometría de un único destino requieren
exactamente una forma seleccionada. El formato de Inicio y la eliminación actúan de forma atómica sobre
varias formas seleccionadas; si el origen protege cualquiera de los destinos, se rechaza todo el lote.

## Crear un dibujo en blanco {#create-a-blank-drawing}

Todo manejador montado o nativo y el elemento compartido exponen
`createBlankDrawing(options?: CreateVsdxOptions): Promise<void>`. El controlador
ofrece el mismo método. Los valores opcionales `width` y `height` usan pulgadas físicas;
por defecto son 8,5 por 11 con una página vacía y una escala de dibujo de 1:1. Archivo > Nuevo >
Dibujo en blanco y Ctrl+N llaman al mismo ciclo de vida. Los controles editables conservan su
comportamiento nativo de atajos.

Una solicitud de Nuevo aceptada reemplaza el documento, la selección y el historial, y el
elemento lo denomina `New drawing.vsdx`. Los bytes de origen se pueden exportar de inmediato;
`dirty` registra las ediciones posteriores. Las cargas reemplazadas, la cancelación y la destrucción
no pueden aceptar un resultado tardío. Los cambios de estado fluyen por la API de estado reactivo
que ya tiene cada enlace.

## Movimiento con el puntero y selección por recuadro {#pointer-movement-and-marquee-selection}

Arrastrar una forma 2D local admisible mueve la selección actual con una traslación común.
Un umbral de cuatro píxeles de cliente preserva los clics normales. La vista previa
solo cambia la presentación SVG; al soltar se envía una única edición atómica del origen. Escape,
la pérdida de la captura del puntero, los cambios de origen, página o selección y los cambios de herramienta la cancelan.
El movimiento usa coordenadas físicas de página y las convierte a unidades de dibujo en el núcleo,
conservando la rotación guardada y los pines locales. Las protecciones del origen siguen siendo la autoridad.

Arrastrar sobre un espacio en blanco de la página selecciona las formas visibles de primer plano
totalmente contenidas, según sus extensiones transformadas de ancho y alto. Mayús, Ctrl o Meta añade a la
selección actual. El movimiento con el puntero excluye por ahora plantillas maestras, grupos, capas, objetos
externos y conectores pegados. Los tiradores de cambio de tamaño y el enrutamiento de conectores son
capacidades distintas; los tiradores de cambio de tamaño se describen más abajo, mientras que el enrutamiento queda fuera
del alcance de este gesto.

## Tiradores de cambio de tamaño y Tamaño y posición {#resize-handles-and-size-position}

La herramienta de puntero compartida muestra ocho tiradores de cambio de tamaño para una forma 2D local
admisible. Los tiradores de borde cambian una dimensión; los de esquina cambian ambas. El
borde o la esquina opuestos permanecen fijos en los ejes locales girados o reflejados que se guardaron para la forma.
Los pines locales descentrados guardados se escalan proporcionalmente cuando la prueba en el origen lo permite.
Durante el arrastre, solo el marco de selección previsualiza los nuevos límites; al soltar
se envía una edición del origen y un paso de deshacer. Cruzar el tirador opuesto limita
la dimensión al menor valor entre su tamaño original y 1/16 de pulgada física, sin
cambiar los indicadores de volteo. El ajuste a la cuadrícula, los gestos con modificador de proporción, la vista previa
del contenido completo y el cambio de tamaño de varias formas no están implementados.

Vista > Paneles de tareas > Tamaño y posición abre el panel numérico compartido. X/Y son el
pin de rotación en pulgadas de dibujo desde el origen inferior izquierdo, con Y creciente
hacia arriba. El ancho y el alto también usan pulgadas de dibujo, independientes de la escala física de la página.
El ángulo usa grados, positivos en sentido antihorario. Intro o salir de un campo
confirma ese campo una sola vez; Escape restaura su valor guardado. Los cambios de ancho y alto
en este panel mantienen fijo el pin de rotación, a diferencia de los tiradores anclados del puntero.
Introducir exactamente el valor actual no cambia los bytes del origen, las fórmulas ni el historial.
Los borradores sobreviven al zoom y a otras ediciones sobre el mismo origen y selección; el reemplazo o
una selección distinta los borra.

Ambos controles requieren un VSDX respaldado por el origen y una única forma 2D local ordinaria, sin capa,
sin pegado y con pines guardados. Los grupos, las plantillas maestras, los objetos externos y la selección múltiple
no están disponibles. La admisibilidad en la escena es preliminar: el núcleo comprueba la geometría del origen,
las protecciones, las fórmulas y las dependencias al enviar. Un rechazo
muestra un error sin cambiar los bytes del origen ni el historial. La cancelación, la destrucción
y una intención más reciente de origen, página o selección impiden que se acepte una edición pendiente.

El asistente del núcleo `visioSizePositionState(page, shapeId)` proporciona un estado numérico
congelado. `visioSizePositionCommand(page, shapeId, field, value)` devuelve ediciones
ordinarias, una matriz vacía para una entrada que no cambia nada, o `undefined` para un candidato
no válido o no admitido. Los controles personalizados pueden usar estos asistentes con la
API de edición atómica existente. El panel numérico no admite cadenas con unidades, entrada de fórmulas
ni cambiar la posición del pin.

## Buscar texto del diagrama {#find-diagram-text}

Los controles de búsqueda compartidos encuentran texto literal en las formas y páginas visibles.
Los anfitriones programáticos usan `controller.setSearchQuery(query)`, `nextSearchResult()`,
`previousSearchResult()` y `selectSearchResult(index)`; la consulta y los resultados
inmutables están disponibles en `controller.state.search`. Consultar deja la
selección intacta hasta que se navega de forma explícita. Un documento nuevo borra la búsqueda, y
las acciones reentrantes más recientes del anfitrión reemplazan la navegación de resultados anterior.

La búsqueda devuelve un resultado por cada forma coincidente, con una vista previa de texto acotada. Usa
la comparación en minúsculas de ECMAScript, no expresiones regulares ni una ordenación sensible al idioma.
Los límites son 256 caracteres de consulta, 500 formas coincidentes, 32 768 caracteres indexados
por forma y 2 millones de caracteres indexados en total. El estado identifica de forma explícita
los resultados parciales. Se excluyen los subárboles ocultos y el texto suprimido de los grupos;
las páginas de fondo se indexan una sola vez como páginas independientes. El resaltado
por coincidencia y la búsqueda en Shape Data no están implementados.

## Capas de visualización {#display-layers}

Use los controles de Capas compartidos o `setLayerVisibility(pageId, layerId, visible)` de cualquier
manejador montado o nativo para anular la visibilidad de visualización.
`visible: null` restaura un indicador guardado; `resetLayerVisibility(pageId?)` restablece
una página de origen o el documento completo. Los ID de página de origen distinguen las capas de primer plano
y de fondo. Las anulaciones congeladas están disponibles en
`controller.state.layerVisibilityOverrides`.

Los cambios de capa sincronizan el renderizado, la selección y la búsqueda sin modificar el
documento. Las guías, NoShow y las formas ocultas heredadas descritas de forma incompleta no pueden
mostrarse mediante una anulación de capa. El panel accesible muestra como máximo 200 capas únicas
y anuncia el truncamiento; la API permite como máximo 25 000 anulaciones con alcance de documento.
El SVG y los artefactos de impresión siguen usando la visibilidad guardada. Los colores de capa, su edición y la
política de impresión guardada no están implementados.

## Pestañas de complementos de la cinta {#ribbon-add-in-tabs}

Un anfitrión puede añadir pestañas después de Ayuda, como hace un complemento en Visio. `ribbonAddIns` es una
propiedad del elemento y una prop de cada enlace de framework, y
`office-ribbon-add-in` forma parte del mapa de eventos:

```tsx
const tabs = [
	{
		id: 'reports',
		label: 'Reports',
		groups: [
			{
				label: 'Export',
				commands: [{ id: 'export', label: 'Export', icon: 'save', run: () => exportReport() }],
			},
		],
	},
];

<VisioViewer
	ribbonAddIns={tabs}
	events={{ 'office-ribbon-add-in': ({ tab, command }) => track(tab, command) }}
/>;
```

Vue usa `:ribbon-add-ins` y emite `office-ribbon-add-in`; Angular usa
`[ribbonAddIns]` y tiene la salida `(ribbonAddIn)`; el enlace vanilla usa
`mountViewer(host, { ribbonAddIns })` y `viewer.update({ ribbonAddIns })`. En
el propio elemento, asigne `element.ribbonAddIns` y escuche el evento del DOM.

Un comando ejecuta su callback `run` y después emite el evento. Si se omite la prop,
se conservan las pestañas actuales; una matriz vacía las elimina. Pasar de nuevo un descriptor
igual, como hace un framework en cada renderizado, conserva los paneles y
solo reemplaza los callbacks. El descriptor (`RibbonAddInTab`, de
`ooxml-core/ribbon`, que también exportan `ooxml-ui` y todos los paquetes de framework) se
comparte con los editores de Word, Excel y PowerPoint. Una pestaña no puede usar el id de
una pestaña integrada (`home`, `insert`, `design`, `data`, `process`, `review`, `view`,
`help`).

## Personalización {#customization}

`customization` acepta un objeto `VisioCustomization`, el contrato compartido de `ooxml-core/embedding` que también usan los demás editores. Cada adaptador lo reenvía como la prop o entrada `customization`. Se aplica cuando cambia la identidad del objeto, así que manténgalo estable; `null` restaura los valores predeterminados.

```ts
const customization = {
	capabilities: { export: false, print: false, copy: false },
	backstage: { extraPages: [{ id: 'policies', label: 'Policies', render: 'acme-policies' }] },
	contextMenu: {
		extraCommands: [{ id: 'send', label: 'Send to chat', onSelect: ({ shapes }) => send(shapes) }],
	},
};
```

- `capabilities`: `false` elimina todas las vías a la capacidad (`edit`, `open`, `save`, `export`, `print`, `share`, `settings`, `copy`, `comment`, `find`): cinta, páginas de la pestaña Archivo, menús y atajos. `hiddenActions` no cambia; ambos se combinan.
- `backstage.extraPages`: páginas del host `{ id, label, group?, order?, render }` en la pestaña Archivo, tras las páginas integradas de su grupo. `render` es una función de montaje, una etiqueta de elemento personalizado o `{ element, properties }` y recibe un `VisioRegionContext`.
- `contextMenu.extraCommands`: comandos del host `{ id, label, icon?, group?, when?, disabled?, onSelect }` encima (`top`, predeterminado) o debajo (`bottom`) de las entradas integradas. `onSelect` se ejecuta tras cerrarse el menú con un `VisioMenuContext`. Las etiquetas se muestran tal cual (localícelas usted mismo).

`edit: false` es un bloqueo de solo lectura propio, junto al interruptor de solo lectura del host y un rol de colaboración `viewer`: ninguno anula a otro. Las capacidades son una política de interfaz, no protección del documento. El reemplazo de regiones (cuadro de opciones, paneles) aún no se ofrece para Visio. Todos los productos juntos: [guía de integración](https://github.com/ChristopherVR/ooxml/blob/main/docs/embedding.md).

## Exportación SVG estática {#static-svg-export}

`exportPageSvg(model, pageIndex, { maxBytes })` y el `exportSvg()` de cada manejador montado o nativo
devuelven una instantánea de la página actual sin modificar el documento,
la selección ni la ventana de visualización. La exportación incluye los fondos admitidos, los recursos
rasterizados incrustados y las notas de compatibilidad visibles en la descripción y los metadatos del SVG. La exportación SVG
no tiene recursos remotos ni descargas automáticas. La demostración solo descarga tras pulsar
un botón de forma explícita.

El límite en UTF-8 es de 16 MiB; quien llama puede reducirlo. Las comprobaciones preliminares prudentes
pueden rechazar contenido cuya serialización final sería menor. Las fuentes no se incrustan y
el texto y la maquetación siguen siendo aproximados. Se conservan las dimensiones físicas de la página;
los rasterizadores posteriores deben limitar las dimensiones y el área en píxeles de su salida
antes de reservar una superficie de imagen. El PDF, la exportación a mapa de bits y la maquetación de impresión siguen
sin implementarse.

Las pruebas secundarias opcionales de píxeles en Linux se ejecutan con
`node scripts/test-svg-rasterization.mjs` tras compilar, con Python 3 del sistema,
librsvg/Cairo y la dependencia de canvas nativo del núcleo. Estas comprobaciones son independientes
de la validación en el navegador o con Microsoft Visio.

## Preparar artefactos de páginas de dibujo {#prepare-drawing-page-artifacts}

`createPrintSnapshot(model, { pageIndices: [0, 2] })` conserva el orden explícito de las páginas
y devuelve registros SVG independientes y profundamente congelados, con las dimensiones del dibujo y
todos los diagnósticos de exportación. El `createPrintSnapshot()` de cada manejador montado o nativo
captura su página actual. Rechaza el reemplazo del documento durante la preparación
sin modificar el estado de la vista. No hay marcos, cuadros de diálogo, descargas ni comandos
de impresora.

La instantánea usa la apariencia de visualización guardada. Las opciones Imprimir y NonPrinting de las capas y la
configuración de impresora guardada no se aplican; las dimensiones del dibujo no son el papel de la impresora ni
la escala de impresión. El resultado incluye estas limitaciones. Las protecciones acotan las páginas seleccionadas,
los bytes agregados, los fondos repetidos, las dimensiones de los rásteres, el trabajo de texto y de trazados y
las validaciones repetidas del modelo completo. Quien llama puede reducir los límites. Los SVG independientes deben
permanecer aislados porque sus ID de recursos internos pueden repetirse. Esto es una
preparación para un futuro flujo de impresión, no compatibilidad con la impresión nativa de Visio.

## Buscar y reemplazar con alcance {#scoped-find-and-replace}

Inicio > Buscar > Reemplazar y Ctrl+H abren los controles de reemplazo compartidos. Buscar siguiente,
Anterior, Reemplazar y Reemplazar todo usan apariciones completas, literales y sensibles a mayúsculas y minúsculas.
Los alcances son Selección, Página actual y Todas las páginas. Selección incluye los
subárboles de los grupos seleccionados en el orden de selección o de origen, eliminando los duplicados de destinos
superpuestos. Los alcances de página usan el orden de origen guardado, incluido el texto de origen oculto.
La navegación da la vuelta y selecciona un destino visible cuando es posible. Coincidir mayúsculas está
marcado y por ahora es fijo; la palabra completa, el plegado de mayúsculas y los códigos de búsqueda especiales
no están implementados. Estas reglas de alcance y recorrido son una política acotada de la aplicación,
no una afirmación de equivalencia con el cuadro de diálogo Buscar nativo.

Todo enlace nativo expone el mismo `controller`: los anfitriones pueden usar
`captureTextReplaceToken(scope)`, `getTextReplaceOccurrences(token, query)`,
`selectTextReplaceOccurrence(token, occurrence, query)`,
`planTextReplacement(token, { query, replacement, mode, current? })` y
`applyTextReplacePlan(plan, token)`. La navegación y la aplicación devuelven un nuevo
token opaco. Conserve ese token para mantener el alcance seleccionado original mientras se pasa
de una coincidencia a otra. Los cambios manuales de selección o de página, el reemplazo del origen u otra
operación invalidan los tokens anteriores. `cancelTextReplace(token)` cancela únicamente la aplicación
activa de ese token, dejando en ejecución otras ediciones más recientes no relacionadas. Los planes y los tokens pertenecen a su controlador
y no pueden falsificarse ni reutilizarse entre instancias.

El reemplazo usa una única edición atómica del origen y un solo borde de deshacer, incluida la navegación
de resultados entre páginas. Los eventos nativos `document-change`, `selection-change` y `shape-select`
siguen disponibles mediante los callbacks del framework. Reemplazar todo conserva la
selección; el reemplazo actual avanza a la siguiente aparición. Un reemplazo con el mismo texto
conserva los bytes exactos del origen y no añade historial, aunque sí navega.
Si el texto coincidente es enriquecido, contiene campos, está vinculado a una plantilla maestra, está protegido o es dependiente,
y no se admite, se rechaza todo el lote. El borrador muestra el error y sigue siendo editable.
Cerrar los controles o cambiar de intención cancela la operación pendiente.

## Edición local experimental de texto sin formato {#experimental-local-plain-text-editing}

Tras `load(bytesOrBlob)`, seleccione una forma local y use los controles de texto compartidos,
o llame a `replacePlainText(pageId, shapeId, text)` en cualquier manejador montado o nativo.
`undo()` y `redo()` usan el historial acotado del documento; `cancelEdit()` cancela
el trabajo pendiente. Los seis adaptadores reenvían los mismos métodos y el mismo evento
`document-change`, cuyo detalle es `{ document, dirty, kind }` (kind `edit`, `undo` o `redo`, más `remote` y `save`; véase arriba).
Inspeccione `controller.state.edit` para conocer la disponibilidad del origen, el estado de ocupado o modificado,
la disponibilidad de deshacer y rehacer, el truncamiento del historial, los errores y los diagnósticos.

Los documentos VSDX respaldados por el origen que el visor carga o crea se pueden editar o
exportar como VSDX.
Las importaciones binarias de VSD v11 usan el códec ole2 publicado a través de ooxml-core: las transformaciones
explícitas de página y forma admitidas, la geometría de movimiento y línea y el texto sin formato se previsualizan
con estilos alternativos diagnosticados. Los bytes heredados originales se conservan de forma privada;
la edición, deshacer y rehacer y la exportación a VSDX están desactivados para los orígenes heredados. Las versiones
binarias anteriores, los grupos y plantillas maestras, la geometría compleja, los stencils, las plantillas y los
formatos con macros siguen sin admitirse. Asignar un modelo con `document` no proporciona
bytes de paquete editables. La validación del núcleo rechaza las formas vinculadas a plantillas maestras, el texto enriquecido,
los campos, los paquetes firmados y el contenido con macros. Las ediciones de XML y del paquete se ejecutan en un
worker aislado; los envoltorios de frameworks no implementan lógica de formato. Las ediciones de texto sin formato
rechazan las cachés de fórmulas dependientes del texto afectadas o desconocidas en lugar de dejarlas
obsoletas, y respetan la protección efectiva `LockTextEdit`. El texto lógico de VSDX
excluye su marcador de párrafo final almacenado, pero conserva los LF finales intencionados
y los párrafos en blanco. La semántica del texto de VSD heredado no cambia. La edición del contenido de texto enriquecido
sigue sin admitirse.

`exportVsdx()` devuelve `{ bytes, dirty, diagnostics }` para una copia descargada de forma explícita.
No sobrescribe el archivo de origen, no lo sube ni afirma fidelidad nativa de ida y vuelta.
La interfaz compartida solo descarga tras una acción explícita del usuario. El historial
está acotado y puede descartar estados de deshacer antiguos, lo que indica `historyTruncated`.
El reemplazo de texto sin formato sigue limitado al subconjunto admitido descrito arriba. Las
operaciones de formato y geometría que figuran a continuación tienen reglas de admisión propias. La evidencia de
reapertura nativa cubre casos registrados concretos, descritos en el
[registro de verificación](/verification); no establece una fidelidad general.

## Formato y organización experimentales {#experimental-formatting-and-arrangement}

La cinta Inicio y `applyEdits()` comparten los comandos `format-text`, `format-shape` y
`reorder-shape`. El formato de texto incluye familias de fuentes guardadas, tamaño en puntos,
negrita, cursiva, subrayado y tachado, color, viñetas, sangría y
alineación horizontal y vertical. El formato de forma incluye relleno sólido y color y grosor de
línea, patrones de línea 0-23, patrones de relleno 0-24, color de fondo y
transparencia. El formato requiere formas hoja locales ordinarias. El formato de texto de forma completa
actualiza cada fila efectiva de carácter o de párrafo, conservando los marcadores de
tramo, el texto, los bits de estilo no relacionados y los campos en caché admitidos. El formato de carácter
y de párrafo tienen admisión de origen independiente; la alineación vertical no reescribe
ninguna de las dos secciones. Las filas protegidas o ambiguas y las dependencias de fórmulas
afectadas rechazan la transacción completa. Aumentar y reducir el tamaño de fuente
requieren un tamaño uniforme; el selector de tamaño puede fijar una selección de tamaños mixtos.
Las celdas locales Character/Paragraph `F="Inh"` son editables cuando se puede demostrar su antecesor y
el valor o las unidades en caché. La delegación sin caché requiere un antecesor real; una elección
explícita de formato crea una anulación local aunque el valor no cambie. Las celdas heredadas
de protección y de transformación de nivel superior siguen protegidas.
Las fuentes ya deben existir en el dibujo de origen. Los comandos de orden de apilamiento admiten
traer al frente, enviar al fondo, avanzar y retroceder en una única forma ordinaria de nivel superior de la banda de visualización cero.

`format-shape` acepta `linePattern`, `fillPattern`, `fillBackgroundColor`
(`'#RRGGBB'`), `lineTransparency` y `fillTransparency` (porcentaje de 0 a 100).
La transparencia se redondea a pasos de medio por ciento, con los empates hacia arriba. La transparencia del relleno
fija a la vez el primer plano y el fondo. El patrón 0 significa sin pintura; el patrón 1 es
sólido. Un `linePattern: 1` explícito activa un contorno con independencia del color de línea.
Un patrón de relleno explícito anula el patrón sólido que implica `fillColor`;
`fillColor: 'none'` con un patrón de relleno distinto de cero se rechaza. La transparencia de un degradado activo
requiere un reemplazo explícito de la pintura, mientras que las pinturas de tema sólidas heredadas
ordinarias se admiten.

El menú Línea compartido ofrece Sin línea y los patrones integrados. Los menús de pintura,
con Más opciones, y el menú contextual de la forma abren el panel de tareas Formato de forma, que
aplica cada cambio al instante como una transacción de selección. Los valores de origen mixtos
aparecen vacíos hasta que se cambian. `fillColorTheme`, `lineColorTheme` y (en
`format-text`) `fontColorTheme` indican la muestra de Colores del tema de la que procede un color
para que se guarde como la fórmula de tema de Visio; `beginArrow`, `endArrow`,
`beginArrowSize`, `endArrowSize`, `lineCap` y `rounding` establecen los extremos de línea,
el remate y el redondeo de esquinas. Los seis enlaces nativos usan el `applyEdits()` existente y
los eventos reactivos de documento y selección; no se añade ningún motor de pintura específico de un framework.
`visioShapeFormattingState(shapes)` proporciona los valores agregados de pintura del origen para
controles personalizados. Estos valores son independientes de los multiplicadores de opacidad de degradados y mosaicos.

Seis comandos de Alinear alinean respecto a la primera forma seleccionada. Posición proporciona un espaciado
igual de bordes horizontal y vertical para tres o más formas. El asistente del núcleo
`visioArrangeCommands(page, selectedIds, action)` produce ediciones atómicas
para los anfitriones con controles personalizados. Mide los cuadros de ancho y alto girados, ordena
la distribución por centros y conserva el orden de selección en los empates. Los grupos, las plantillas maestras,
las capas y los conectores pegados quedan fuera de su ámbito de admisión. Estas acciones
comparten con las demás ediciones el historial, la conservación de la selección y la descarga y reapertura de VSDX.

## Duplicación experimental {#experimental-duplication}

`duplicateSelection(): Promise<void>` en el controlador y en todo manejador montado o nativo
duplica las formas seleccionadas admitidas de la página actual. El menú Pegar de Inicio,
el menú contextual de la forma, Tell me y Ctrl+D llaman al mismo método. Los campos
de texto conservan su comportamiento nativo de teclado. Duplicar no usa el portapapeles
del sistema.

El asistente del núcleo `visioDuplicateCommand(page, selectedIds)` produce un
comando de origen `duplicate-shapes` con ID nuevos locales a la página. Las copias conservan el XML de origen,
el texto mixto y el orden de apilamiento, reciben nombres nuevos sin los ID únicos copiados,
y se desplazan 0,33 pulgadas de dibujo a la derecha y hacia abajo. La validación del origen admite por ahora
formas hoja 2D locales y sin pegar, con exclusión de plantillas maestras, grupos, contenido externo,
capas, identidades ambiguas y dependencias no admitidas. La protección de pines y las cachés
de fórmulas afectadas se comprueban antes de aceptar cualquier cambio en el origen.

Las copias resultantes pasan a ser la selección en un único paso atómico de historial. Deshacer
restaura la selección original y rehacer las copias, salvo que una selección o una acción de página más reciente lo reemplace.
Las operaciones fallidas, canceladas o reemplazadas no pueden cambiar parcialmente los bytes del origen ni la selección.
Los eventos ordinarios `selection-change` y `document-change` exponen el resultado aceptado al
estado nativo del framework.

## Portapapeles de formas experimental {#experimental-shape-clipboard}

Todo manejador montado o nativo expone `copySelection()`, `cutSelection()` y
`pasteSelection()`, y cada uno devuelve `Promise<void>`. Invóquelos desde un gesto
del usuario para que el navegador pueda conceder el acceso al portapapeles. Inicio, los menús contextuales, Tell me
y Ctrl/Meta+C, X y V usan el mismo adaptador compartido. Los campos de texto conservan el
comportamiento nativo del portapapeles. Los fallos del portapapeles aparecen en el visor sin cambiar los
bytes del origen, el historial ni la selección. Cortar escribe en el portapapeles antes de eliminar;
si la protección impide eliminar, las formas copiadas permanecen en el portapapeles y
el visor informa de que no se pudieron cortar.

La carga útil es `text/plain` acotado con el marcador `OOXML-VISIO-SHAPES/1`. Contiene
el XML de las formas capturadas y el contexto de origen necesario, de modo que editar o cortar
los originales no invalida el contenido copiado. Pegar lee siempre el portapapeles actual
del sistema; no reutiliza en silencio una copia interna anterior.
Los formatos nativos de portapapeles de Microsoft Visio, el texto arbitrario, las imágenes y Pegado
especial no se admiten. Las formas pegadas reciben ID y nombres nuevos y un desplazamiento
de 0,33 pulgadas de dibujo a la derecha y hacia abajo. Pegar de nuevo usa ese mismo desplazamiento;
la cascada de pegado nativa y la ubicación entre páginas no están establecidas.

El primer alcance admitido son las hojas 2D locales ordinarias sin pegar, con definiciones coincidentes de
fuente, estilo y tema del documento y con configuración de página, ordinal, contexto de fondo
y proporción de dibujo equivalentes. Las referencias deben permanecer dentro de la selección copiada.
Los grupos, las plantillas maestras, las capas, el contenido externo, los fragmentos con relaciones,
las fórmulas no admitidas y la importación de recursos siguen rechazándose. El núcleo valida
toda la carga útil antes de una edición atómica del origen. Pegar selecciona sus copias y comparte
el historial acotado y las reglas de intención de selección de Duplicar.

`controller.state.clipboard` expone el estado inmutable `ready`, `preparing` y `error`
mediante la suscripción reactiva que ya tiene cada framework. La captura del origen
se ejecuta en un worker cancelable y se invalida por cambios de selección, página, visibilidad
o documento. Los anfitriones que aportan su propio transporte de portapapeles pueden usar
`captureClipboardToken()`, `prepareClipboardSelection(token)`,
`getPreparedClipboard(token)`, `cutPreparedSelection(token)` y
`pasteClipboardText(text, token)` en el controlador. Los tokens rechazan acciones obsoletas.
El núcleo sin DOM exporta `captureVisioClipboard`, `serializeVisioClipboard`,
`deserializeVisioClipboard`, la edición `paste-shapes` y el constructor de comandos
`visioPasteCommand(page, clipboard, offset?)`.

## Edición de geometría experimental {#experimental-geometry-editing}

Los controles de geometría compartidos exponen crear rectángulo, mover, cambiar de tamaño y eliminar
de forma segura. Todo manejador de framework nativo o montado reenvía además
`applyEdits(edits: readonly VisioEdit[])` para un lote atómico por el mismo
camino de worker, historial y cancelación.

Inicio > Herramienta de texto e Insertar > Cuadro de texto usan el mismo borrador compartido de cuadro de texto
de tamaño fijo. Arrastre sobre un espacio en blanco de la página, escriba texto sin formato y después use Añadir cuadro de texto o
Ctrl/Meta+Intro; Escape lo cancela. Los borradores vacíos no crean formas de origen;
se conservan los espacios y saltos de línea intencionados. Al hacer clic en una forma existente con la Herramienta de texto, se selecciona y se muestran
los controles de edición de texto existentes. El dibujo de rectángulos, elipses y líneas usa el
mismo ciclo de vida de creación protegido. Su transacción interna del controlador captura la intención
de origen, página, selección ordenada, zoom y visibilidad antes de dibujar,
comprueba esa intención antes de aceptar el origen y selecciona las formas creadas en el
mismo paso de deshacer. Una intención más reciente o la destrucción reemplazan una creación pendiente. Los tokens
internos de creación no son una API adicional de los enlaces nativos.

Los seis manejadores públicos aceptan el nuevo comando ordinario mediante `applyEdits()`:

```ts
await handle.applyEdits([
	{
		type: 'create-text-box',
		pageId: '0',
		shapeId: '10',
		x: 4,
		y: 5,
		width: 3,
		height: 2,
		text: 'Text with an intentional trailing blank paragraph\n',
	},
]);
```

El texto es obligatorio. El estilo de texto predeterminado guardado aporta el formato inicial de los
caracteres. `FillPattern=0` y `LinePattern=0` explícitos dejan el nuevo cuadro sin
pintura; cambiar solo el color de línea no activa un contorno. La creación de un borrador de tamaño fijo
no establece los valores predeterminados de la Herramienta de texto de la interfaz nativa, el dimensionado automático del texto
ni la edición de contenido de texto enriquecido. `applyEdits()` genérico conserva su comportamiento
de selección existente; la herramienta de dibujo compartida es la que gestiona la selección de la forma creada.
La evidencia nativa cubre 24 casos de API de cuadro fijo guardados con tres escalas de dibujo,
con estilos predeterminados y personalizados de Solo texto/Arial 18 pt, y con las salidas del núcleo reabiertas
en Visio. Estos casos de estilo guardado no miden los valores predeterminados interactivos de la interfaz; consulte el
[registro de verificación](/verification).

Insertar > Página nueva se encamina por el comando existente de inserción de página
respaldado por el origen y después selecciona la página nueva. Comparte el comportamiento de inserción
de la barra de páginas, deshacer y rehacer y la admisión del origen; esto no añade la creación de páginas basada en plantillas.

```ts
await handle.applyEdits([
	{
		type: 'create-rectangle',
		pageId: '0',
		shapeId: '9',
		x: 4,
		y: 5,
		width: 3,
		height: 2,
		text: 'New shape',
	},
	{ type: 'move-shape', pageId: '0', shapeId: '1', x: 6, y: 7 },
]);
```

Las coordenadas son posiciones del pin de rotación en pulgadas de dibujo, con origen inferior izquierdo
y positivo hacia arriba. El cambio de tamaño mantiene fijo el pin salvo que el comando incluya
`anchor: { x, y }`. Cada coordenada de anclaje es `0`, `0.5` o `1`, e identifica el
punto fijo en los límites locales normalizados de ancho y alto de la forma, orientados hacia arriba. Por
ejemplo, `{ x: 0, y: 0 }` mantiene fija su esquina inferior izquierda local. El cambio de tamaño anclado
tiene una prueba de origen independiente para los pines locales, las cachés afectadas y los valores fijos
de rotación y volteo. Los ID son explícitos. Las formas 2D locales de nivel superior ya admitidas
se pueden editar, incluidas las importadas de otros productores. El núcleo evalúa las dependencias numéricas de ShapeSheet
afectadas que se admiten y conserva las cargas útiles ZIP no tocadas. El visor nunca modifica el XML por sí mismo.

Este es un subconjunto admitido y estrecho. Las plantillas maestras, los grupos y las formas externas,
los conectores y el pegado, las protecciones o redirecciones inseguras, las fórmulas afectadas no admitidas,
las dependencias ambiguas del paquete y la eliminación de elementos referenciados fallan con un
error visible. La geometría de línea relativa se escala; la geometría de línea absoluta necesita una
dependencia de dimensión admitida. Consulte el
[`src/core/visio/README.md`](https://github.com/ChristopherVR/ooxml/blob/main/src/core/visio/README.md) del núcleo
para ver el contrato de los comandos, los límites numéricos, los motivos de exclusión y las próximas ampliaciones.
El [registro de capacidades](/parity) recoge el conjunto admitido actual, que ahora
incluye también los movimientos acotados de instancias primitivas de plantillas maestras con pines locales explícitos.

Las pruebas de formas sencillas generadas o importadas y un flujo de trabajo local en Chrome pasan. Los 19
diagramas del corpus público aceptados admiten actualmente cero ediciones de geometría porque
todavía no se puede demostrar que el alcance de dependencias ajenas a la página sea independiente. Trece carecen además
de una forma local admisible. La cobertura práctica de edición para ese corpus sigue bloqueada;
el siguiente trabajo es un grafo de dependencias acotado de paquete, plantilla maestra y tema, y la edición
de instancias de plantillas maestras, seguidos del enrutamiento de extremos pegados. Desde entonces se ha verificado la reapertura nativa
para las capturas concretas de formato y organización del
[registro de verificación](/verification); la fidelidad general sigue sin verificarse.
