# Colaboración

Archivo > Compartir inicia una sesión en vivo entre ventanas de un mismo navegador. Compartir entre dispositivos aún no está disponible.

::: warning Parcial, solo en el mismo navegador
En esta compilación no hay ningún servicio alojado ni transporte de servidor. Nada sale de su navegador y no se sube ningún documento.
:::

## Qué existe hoy {#what-exists-today}

- Una sesión en vivo basada en `ooxml-core/collab` (Yjs) sobre un `BroadcastChannel`, de modo que las ventanas y pestañas del mismo navegador pueden compartir un dibujo.
- Se comparte todo el paquete VSDX, y el entorno de ejecución de colaboración se carga solo cuando se inicia el uso compartido.
- Los cambios remotos llegan como un único paso de historial que se puede deshacer, y los colaboradores aparecen mediante el elemento de presencia compartido de `ooxml-ui`.
- El nombre de la sesión tiene de 1 a 64 letras, dígitos, guiones o guiones bajos. Una ventana que se une a una sesión cuya sala ya contiene un dibujo lo adopta; en caso contrario, inicializa la sala.

## Demostración en vivo: dos ventanas, cualquier framework {#live-demo-two-windows-any-framework}

Una sesión de uso compartido es solo un nombre. Los dos paneles de abajo son áreas de pruebas independientes, creadas con enlaces de frameworks distintos, que se unen a la sesión `docs-collab` al cargarse: la ventana A es el área de pruebas vanilla y abre el ejemplo; la ventana B es el área de pruebas de React y recibe el dibujo de la sala. Edite el texto de una forma en un panel y vea cómo el otro lo sigue; deshacer revierte el paso remoto.

<iframe
	src="/ooxml/visio/demo/?embed=1&sample=1&share=docs-collab"
	title="visio-viewer window A (vanilla)"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

<iframe
	src="/ooxml/visio/demo-react/?embed=1&share=docs-collab"
	title="visio-viewer window B (React)"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px; margin-top: 1rem"
></iframe>

### Combine dos frameworks cualesquiera {#pair-any-two-frameworks}

Todas las áreas de pruebas aceptan `?share=<session>`, que rellena Archivo > Compartir y lo inicia. Solo importa el nombre de la sesión, así que sirve cualquier combinación: la [demostración en vivo](/es/#live-demo) de la página de inicio tiene un selector de framework para cada ventana, y en dos pestañas puede abrir, por ejemplo, `/demo-vue/?sample=1&share=my-room` y `/demo-angular/?share=my-room`. También puede hacerlo a mano en dos ventanas cualesquiera: Archivo, luego Compartir, y el mismo nombre de sesión.

### Por qué solo funciona en el mismo navegador {#why-it-is-same-browser-only}

La sesión usa `ooxml-core/collab` (Yjs) sobre un `BroadcastChannel`. Un `BroadcastChannel` conecta ventanas, pestañas y marcos de un mismo perfil de navegador y un mismo origen, sea cual sea el framework con el que se hayan creado, y nada más. Dos máquinas distintas no pueden encontrarse así; para eso hace falta un transporte de servidor que aporte la aplicación anfitriona, y esta compilación no incluye ninguno.

## Qué se admite y qué no {#what-is-and-is-not-supported}

| Admitido                                                                           | No admitido                                                                                                       |
| ---------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------- |
| Sesiones en el mismo navegador sobre BroadcastChannel, basadas en Yjs              | Compartir entre dispositivos: requiere un transporte de servidor aportado por el anfitrión, no se incluye ninguno |
| Uso compartido del paquete completo, con un paso reversible por cada cambio remoto | Fusión a nivel de campo: los cambios concurrentes no se fusionan, gana el último paquete                          |
| Colaboradores mostrados con el elemento de presencia compartido                    | Comentarios, autenticación y roles                                                                                |
| Archivo > Compartir como punto de entrada                                          | Una API de incrustación documentada para transportes personalizados                                               |

La evidencia consta en el [registro de capacidades](/parity) bajo "Collaboration and comments".
