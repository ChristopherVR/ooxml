# Colaboración

Varias personas pueden editar un mismo libro al mismo tiempo. El editor asigna el libro a un documento de Yjs mediante `ooxml-core/xlsx/collab` y se une a una sala a través de la sesión compartida `ooxml-core/collab` (la misma sesión y los mismos proveedores que usan los editores de Word y Visio). Cada edición se envía en el momento en que se realiza, las ediciones simultáneas en celdas distintas se fusionan y cada persona ve dónde están las demás.

## Empezar a compartir

Desde el editor: **Archivo > Compartir** (o el botón Compartir en el extremo derecho de la fila de pestañas de la cinta), escriba un nombre de sesión y elija **Iniciar uso compartido**. Otra ventana o pestaña del mismo navegador que tenga un libro abierto e inicie el uso compartido con el mismo nombre de sesión se une a ella. Sin un servidor definido por la página, nada sale del navegador: la sesión se ejecuta sobre un `BroadcastChannel`.

Desde el código, defina la propiedad `collaboration` (o prop, en cada enlace):

```ts
editor.newWorkbook(); // a window joining an existing room still needs a workbook open
editor.collaboration = {
	roomId: 'budget-2026',
	serverUrl: 'wss://collab.example.com/rooms', // y-websocket compatible; omit for same-browser only
	user: { name: 'Ada', color: '#2563eb' },
};
editor.addEventListener('collaboration-change', (event) => console.log(event.detail));
```

| Opción      | Notas                                                                                                                                           |
| ----------- | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| `roomId`    | De 1 a 128 letras, dígitos, guiones o guiones bajos.                                                                                            |
| `serverUrl` | URL base de un servidor compatible con y-websocket; se le añade el identificador de la sala.                                                    |
| `provider`  | Una fábrica de proveedor propia (señalización WebRTC o un proveedor de Yjs existente mediante `externalProvider`). Prevalece sobre `serverUrl`. |
| `session`   | Una `CollabSession` existente creada con `sanitizePayload: sanitizeXlsxPresence`. El editor nunca destruye una sesión que se le ha entregado.   |
| `user`      | `name` (por defecto `authorName`), `color` (por defecto un color de la paleta), `role` (`viewer` pone el editor en modo de solo lectura).       |

Mientras `collaboration` está definida, un libro que se abra después se une a la misma sala; cuando la sala ya tiene un libro, el contenido de la sala sustituye al abierto. Si se establece en `null`, se abandona la sala. Las formas imperativas son `startCollaboration(options)` (se resuelve una vez unido), `stopCollaboration()`, `reconnectCollaboration()`, `collaborationState` y `share()` (abre Archivo > Compartir). Todos los enlaces reenvían `collaboration`, notifican `collaboration-change` (`onCollaborationChange`, en Vue `collaboration-change`, en Angular `collaborationChange`, en Svelte `oncollaborationchange`) y exponen `share()` en su manejador; al desmontar el editor se abandona la sala.

## Qué se comparte

- Valores de celda, fórmulas (con sus valores en caché), formatos numéricos y formato de celda.
- Hojas: añadir, cambiar el nombre, mover, eliminar, ocultar y color de la pestaña.
- Alturas de fila, anchuras de columna, filas y columnas ocultas, celdas combinadas y nombres definidos.
- La selección de cada persona, dibujada en las cuadrículas de las demás como un contorno de su color con su nombre, y las personas de la sesión como avatares en la barra de título.

Deshacer (Ctrl+Z, la cinta y la Barra de herramientas de acceso rápido) solo deshace sus propias ediciones mientras el libro está compartido; el trabajo de un colaborador nunca se revierte. Cuando se pierde la conexión, el editor lo indica y conserva sus ediciones; se envían cuando la conexión vuelve, y un segundo aviso confirma la reconexión.

## Qué no se comparte

Los gráficos, las imágenes y otros dibujos, los comentarios, los hipervínculos, los formatos condicionales, la validación de datos, las tablas y los filtros, las tablas dinámicas, la inmovilización de paneles y otros ajustes de vista, la configuración de página, la protección, las propiedades del documento y los ajustes de cálculo permanecen en cada ventana y se guardan desde ella. Las inserciones y eliminaciones de filas y columnas viajan como las celdas que desplazan, de modo que una edición simultánea de una celda desplazada llega a su dirección anterior. Una fórmula que un colaborador escribió con el nombre de una hoja que se renombró al mismo tiempo conserva el nombre antiguo. Las reglas de conflicto se enumeran en el repositorio `ooxml` (`docs/collab-area.md`).

## Qué corresponde al anfitrión

El editor no incluye ningún servidor. Su aplicación es responsable del transporte (la URL del servidor o el proveedor), de la autenticación y la autorización (los roles son orientativos en el cliente), de la persistencia de las salas y de cuándo y desde qué ventana se guarda el libro: el editor no coordina guardados simultáneos. Esto no es la coautoría de Excel ni afirma paridad con ella.
