# Enlaces de framework

Seis integraciones nativas y ligeras usan un único elemento del navegador y una sola implementación de `mountViewer`. Cada una se publica como paquete independiente; ninguna analiza documentos, dibuja diagramas ni mantiene otro renderizador.

Cada handle nativo expone `createBlankDrawing(options?: CreateVsdxOptions)`.
El controlador compartido crea el nuevo origen y lo acepta a través de su ciclo de carga existente; los hooks de estado y los eventos de documento se actualizan con normalidad. El movimiento del puntero y la selección por rectángulo se ejecutan dentro del mismo componente web compartido para todos los enlaces.

| Framework | Paquete                | Exportación                                  |
| --------- | ---------------------- | -------------------------------------------- |
| Vanilla   | `visio-vanilla-viewer` | `mountViewer`                                |
| React     | `visio-react-viewer`   | `VisioViewer`                                |
| Vue 3     | `visio-vue-viewer`     | `VisioViewer`                                |
| Angular   | `visio-angular-viewer` | `VisioViewerComponent` (`visio-viewer-host`) |
| Svelte 5  | `visio-svelte-viewer`  | `VisioViewer` por defecto                    |
| Solid     | `visio-solid-viewer`   | `VisioViewer`                                |

Importe un paquete de framework específico para evitar cargar entornos de ejecución de otros frameworks. El código fuente de los adaptadores está en `packages/bindings` del repositorio, y su README enumera las comprobaciones que los protegen.

La conexión es compartida. El contrato (`VIEWER_CONTRACT`, las claves de propiedades y eventos y los miembros del handle) y el reenvío de parches con reconocimiento de revisiones de `mountFrameworkViewer` están en `packages/bindings/src/common.ts`. La maquinaria genérica de enlace (el handle diferido, la selección de propiedades y las factorías de componentes de React y Vue) es el kit `ooxml-ui/bindings`. React y Vue se construyen a partir de esas factorías; Angular instala los miembros del handle desde `VIEWER_CONTRACT`; Svelte es un componente; Solid es un componente JSX que monta directamente el elemento compartido. `packages/bindings/tests/contract-parity.test.ts` comprueba cada adaptador frente al contrato: las mismas propiedades, cada evento mediante el mapa de callbacks y la salida nativa, y cada miembro del handle.

## Contrato compartido {#shared-contract}

Todos los adaptadores aceptan las mismas propiedades:

- `document`: un `VisioDocument` o `null` para limpiarlo
- `pageIndex`: índice de página con base cero
- `zoom`: escala, donde `1` equivale al 100 %
- `showToolbar`: visibilidad de la barra de herramientas
- `canEdit`: permiso de edición. El visor completo es editable por defecto; `false` lo deja en solo lectura (la selección, la búsqueda, el zoom y la exportación siguen funcionando). Los bloques de construcción son de solo lectura por defecto, consulte [bloques de construcción](/es/building-blocks)
- `readOnly`: una restricción propia. `true` hace el dibujo de solo lectura y prevalece sobre `canEdit`; `false` u omitido levanta únicamente esa restricción
- `ribbonAddIns`: pestañas que el anfitrión añade después de Ayuda (consulte [Pestañas de complementos de la cinta](/es/api#ribbon-add-in-tabs))
- `collaboration`: comparte el dibujo abierto en una sala mientras esté definido; `null` lo deja (consulte [colaboración](/es/collaboration))
- `storage` y `customization`: guardado y apertura gestionados por el anfitrión, y el objeto de personalización compartido (consulte [Personalización](/es/api#customization) y [Almacenamiento gestionado por el anfitrión](/api#host-owned-storage-and-file-commands))
- `fileProviders`: ubicaciones en la nube para Archivo > Abrir, como OneDrive (consulte [OneDrive](/es/onedrive))
- `events`: mapa de callbacks tipado con nombres de eventos literales

El mapa de eventos contiene actualmente `document-load`, `document-change`, `document-error`, `page-change`, `zoom-change`, `shape-select`, `selection-change`, `office-ribbon-add-in`, `collaboration-change` y `file-command`. Se deriva del contrato canónico. Todos los adaptadores transmiten todos los eventos; ninguno analiza documentos ni mantiene otro renderizador.

Las actualizaciones de propiedades son parciales. Las instantáneas de los frameworks nativos se comparan para que los renderizados no relacionados no deshagan una carga imperativa ni un zoom cambiado por el usuario. Las propiedades omitidas o `undefined` conservan el estado actual del visor. `null` borra explícitamente el documento. Quitar la prop `events` elimina los manejadores del mapa de callbacks. Proporcione un nuevo objeto de documento para una sustitución externa; la mutación en el mismo objeto no es un mecanismo de reactividad admitido. Las interacciones de página y zoom pueden actualizar el estado interno del visor; las props solicitan cambios cuando el framework las actualiza.

`document-change` transporta `{ document, dirty, kind }`, donde kind es `edit`, `undo`, `redo`, `remote` (un cambio de un colaborador) o `save` (tras `markClean()`). El `ViewerHandle` imperativo expone `element`, `controller`, `load(bytesOrBlob)` y `fit()`. Los métodos de edición son `applyEdits(edits)`, `replacePlainText(pageId, shapeId, text)`, `duplicateSelection()`, `copySelection()`, `cutSelection()`, `pasteSelection()`, `undo()`, `redo()`, `cancelEdit()` y `exportVsdx()`. Cuando el anfitrión ha guardado los bytes exportados, llama a `controller.markClean()`. La colaboración se controla mediante la propiedad; `element.collaborationState`, `element.reconnectCollaboration()` y `element.stopCollaboration()` están en el `element` del handle. El estado de preparación del portapapeles está disponible en `controller.state.clipboard` a través de cada suscripción reactiva nativa; los métodos del portapapeles requieren un gesto del usuario en el navegador y admiten el formato de forma de visor acotado. Se transfieren a la misma implementación de origen aislada en un worker. `ViewerEditState` y `VsdxExportResult` son tipos exportados compartidos; el estado está disponible en `controller.state.edit`. Angular expone la salida nativa `documentChange`; todos los adaptadores admiten el evento `document-change` del mapa de callbacks. El acceso de lectura y las operaciones a través de un handle de framework fallan de forma explícita antes del montaje o después del desmontaje. Un controlador o elemento retenido también se destruye cuando su propietario se desmonta. Vanilla devuelve el enlace compartido completo con `update()` y `destroy()` idempotente.

El editor se monta en el cliente. Las pruebas de Node por separado verifican que los seis puntos de entrada se importan sin globales del navegador; React, Vue, Svelte y Solid también renderizan hosts nativos vacíos en el servidor. React y Vue tienen pruebas que hidratan su marcado real del servidor antes de montar el elemento compartido. El renderizado en el servidor de Angular y la hidratación de Solid y Svelte siguen sin verificarse, como el montaje entre documentos. Estas comprobaciones usan las versiones instaladas y no implican una cobertura completa de integración con frameworks de renderizado en el servidor.

## Uso nativo {#native-usage}

React usa una ref reenviada y props habituales:

```tsx
const ref = useRef<ViewerHandle>(null);
<VisioViewer ref={ref} document={diagram} events={{ 'document-error': reportError }} />;
```

Vue expone el mismo handle en la ref de su componente y además emite los eventos nativos en kebab-case:

```vue
<VisioViewer ref="viewer" :document="diagram" @document-error="reportError" />
```

Angular expone entradas y salidas nativas. Añada el componente independiente a los imports de la aplicación:

```html
<visio-viewer-host [document]="diagram" (documentError)="reportError($event)" />
```

Las salidas de Angular son `documentLoad`, `documentChange`, `documentError`, `pageChange`, `zoomChange`, `shapeSelect`, `selectionChange`, `ribbonAddIn`, `collaborationChange` y `fileCommand`. `fileCommand` solo informa del comando; para cancelarlo, un anfitrión de Angular devuelve `false` desde el callback `events['file-command']`. El mapa completo de callbacks `events` también se acepta como entrada. Vue y Angular invocan el manejador del mapa de callbacks y después emiten su salida nativa; use una de las dos APIs, salvo que quiera ambas notificaciones.

Para modelos de documento grandes e inmutables, use `shallowRef`/`markRaw` en Vue y `$state.raw` en Svelte en lugar de crear un proxy profundo de un objeto de opciones nuevo en cada actualización. Envolver un documento en un nuevo proxy reactivo cambia su identidad y solicita una sustitución externa. La mutación en el mismo objeto no es compatible.

Svelte usa props de callback nativas a través de `events`. Su componente exporta `load`, `fit` y `getHandle`, de modo que el handle tipado completo sea accesible sin exponer el DOM privado:

```svelte
<VisioViewer bind:this={viewer} document={diagram}
  events={{ 'document-error': reportError }} />
```

Solid proporciona un callback de ref con el handle al montarse y `undefined` al desmontarse o al reemplazar la ref:

```tsx
<VisioViewer
	document={diagram()}
	viewerRef={setViewer}
	events={{ 'document-error': reportError }}
/>
```

Vanilla solo gestiona el hijo que crea:

```ts
const viewer = mountViewer(host, { document: diagram, zoom: 1 });
viewer.update({ events: { 'document-error': reportError } });
// Later:
viewer.destroy();
```

El estilo es nativo de cada framework: `className`/`style` en React, atributos heredados en Vue, atributos del host en Angular, `class`/`style` en Svelte y `class`/`style` en Solid. La aplicación debe dar al host una altura utilizable. El estilo compartido del visor y la accesibilidad residen en el elemento personalizado.
