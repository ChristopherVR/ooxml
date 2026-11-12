# Enlaces de OneDrive y documentos de Office

Para una configuración paso a paso (registro en Microsoft Entra, inicio de sesión y un ejemplo
completo), empieza por la [guía rápida de OneDrive](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md).

OpenTeams puede contener un enlace de OneDrive a un archivo DOCX, XLSX, PPTX o VSDX. La suite Office
integrada abre ese archivo con el editor de documentos correspondiente y la conexión compartida
`ooxml-core/embedding/onedrive`. El propio OpenTeams es un espacio de trabajo, no un quinto editor
de archivos de Office y no es Microsoft Teams.

La misma distinción se aplica a `openteams-react-viewer`, `openteams-vue-viewer`,
`openteams-angular-viewer`, `openteams-svelte-viewer`, `openteams-solid-viewer`,
`openteams-vanilla-viewer` y al elemento `ooxml-ui/teams`:

| Elemento compartido                          | Dónde se guardan las ediciones                                                                             |
| -------------------------------------------- | ---------------------------------------------------------------------------------------------------------- |
| Enlace de OneDrive                           | En el elemento original de OneDrive, solo cuando el host autorizado guarda mediante su sesión de OneDrive. |
| Enlace a archivo de espacio de trabajo local | En el espacio de trabajo local de Office; no es un enlace de nube entre dispositivos.                      |
| Adjunto del servidor de OpenTeams            | En el almacenamiento de ese servidor; una copia subida es independiente del original de OneDrive.          |

El flujo de Compartir de los archivos conectados de la suite solicita un rol explícito de
visualización o edición y un ámbito de enlace. Publicar o copiar la URL de un elemento existente no
concede por sí mismo acceso a otra persona. Nunca crees un enlace anónimo como alternativa cuando un
tenant rechace un enlace restringido. `openteams-server` no obtiene credenciales de Microsoft, no
replica carpetas de la unidad ni fusiona los cambios de los archivos de Office.

Usa la [guía de integración de OneDrive de Office](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#sharing-and-openteams)
para la autenticación, las referencias de paquete y el uso compartido según permisos. Su
[guía de pruebas](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md) distingue
los transportes de prueba locales de la aceptación con cuentas reales. No se incluye ninguna interfaz
ni servicio de simulación.

Las salas de colaboración de chat/presencia y de documentos de Yjs son independientes de la
coautoría de Microsoft Office. Enviar la misma URL de OneDrive a dos personas no sincroniza
automáticamente sus editores.
