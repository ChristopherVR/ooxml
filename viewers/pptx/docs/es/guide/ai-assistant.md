# Asistente de IA

Cada enlace incluye un panel de chat de IA opcional y activable. El usuario hace clic en el icono de IA (destellos) de la barra de herramientas y se abre un panel de chat en el lado derecho del visualizador. El asistente puede leer la presentación, responder preguntas sobre ella y proponer cambios (texto, estilos, diseño, diapositivas, tema) que el usuario revisa y acepta o rechaza antes de que se modifique nada.

La función es **independiente del proveedor** y **usa su propia autenticación**: la biblioteca no incluye ningún modelo ni clave de API. Usted conecta el asistente al modelo que controle (una ruta de backend, un modelo en el navegador o un transporte totalmente personalizado). Si no pasa una configuración `ai`, no aparece ningún icono y no se carga nada del código de IA (ni el SDK `ai`), así que no hay ningún coste cuando la función no se usa.

## Instalación {#installation}

Los paquetes de IA son **dependencias de pares opcionales**. Instale `ai` y, para React, Vue o Svelte, el paquete de UI del SDK de IA de ese framework. Angular y Vanilla usan el controlador independiente del framework incluido en el visualizador, así que solo necesitan `ai`.

```bash
# React
bun add ai @ai-sdk/react
# Vue
bun add ai @ai-sdk/vue
# Svelte
bun add ai @ai-sdk/svelte
# Angular / Vanilla
bun add ai
```

También instala un **proveedor** para el modelo que utilice. Nunca es una dependencia del visualizador; lo elige usted:

```bash
# Any OpenAI-compatible endpoint (Ollama, vLLM, LM Studio, a proxy, ...)
bun add @ai-sdk/openai-compatible
# or a first-party provider
bun add @ai-sdk/openai
bun add @ai-sdk/anthropic
```

## Activar el panel {#enabling-the-panel}

Pase un asistente a la prop `ai` del visualizador (una entrada en Angular, una opción en vanilla). Cada enlace tiene una entrada `/ai` cuya función `createAiAssistant(config)` convierte una configuración en un controlador de asistente. La forma de la configuración es idéntica en todos los enlaces; solo cambia la sintaxis del host para pasar una prop.

```tsx
// React
import { PowerPointViewer } from 'pptx-react-viewer';
import { createAiAssistant } from 'pptx-react-viewer/ai';

const ai = createAiAssistant({ connection: { kind: 'endpoint', api: '/api/ai/chat' } });

<PowerPointViewer content={deck} ai={ai} />;
```

```ts
// Vue, Angular, Svelte and vanilla: the same factory from the binding's own `/ai` entry
import { createAiAssistant } from 'pptx-vue-viewer/ai';
// or 'pptx-angular-viewer/ai', 'pptx-svelte-viewer/ai', 'pptx-vanilla-viewer/ai'

const ai = createAiAssistant({ connection: { kind: 'endpoint', api: '/api/ai/chat' } });
```

```vue
<!-- Vue -->
<PowerPointViewer :content="deck" :ai="ai" />
```

```html
<!-- Angular -->
<pptx-viewer [content]="deck" [ai]="ai" />
```

```svelte
<!-- Svelte -->
<PowerPointViewer source={deck} {ai} />
```

```ts
// Vanilla
import { createPptxViewer } from 'pptx-vanilla-viewer';

createPptxViewer(container, { source: deck, ai });
```

Pasar la configuración simple en su lugar (`ai: { connection }`) sigue funcionando en todos los enlaces y no está obsoleta en esta versión. Cree el controlador una sola vez (a nivel de módulo, en `useMemo` o como campo de componente); un controlador que se recrea en cada renderización sigue manteniendo montado el chat abierto. Consulte [La entrada `/ai`](#ai-entry) para saber qué garantiza esta versión y qué cambia la siguiente versión principal.

## Modos de conexión {#connection-modes}

`connection` declara cómo llega el asistente a un modelo. Hay tres tipos.

### `endpoint` (recomendado para producción) {#endpoint-recommended-for-production}

Envía los mensajes a una ruta de su propio backend. La clave del proveedor permanece en el servidor; el navegador nunca la ve. Es el valor por defecto más seguro.

```ts
const ai: PptxAiConfig = {
	connection: {
		kind: 'endpoint',
		api: '/api/ai/chat',
		// Headers and body may be functions, so auth tokens can be resolved per
		// request (for example from an OAuth session that refreshes).
		headers: async () => ({ Authorization: `Bearer ${await getAccessToken()}` }),
		credentials: 'include',
	},
};
```

En el servidor, ejecute el modelo con el SDK de IA y exponga las herramientas de la presentación como herramientas **solo de esquema**. Como las herramientas no tienen `execute`, las llamadas del modelo se envían en flujo al navegador, se ejecutan allí sobre la presentación activa y los resultados vuelven en flujo al modelo. Su clave nunca sale del servidor.

```ts
import { openai } from '@ai-sdk/openai';
// A backend route (framework-agnostic helpers from `ooxml-ui/pptx/ai`)
import { streamText } from 'ai';
import { buildPptxAiTools, buildPptxAiSystemPrompt } from 'ooxml-ui/pptx/ai';

export async function POST(request: Request) {
	const { messages } = await request.json();
	const result = streamText({
		model: openai('gpt-4o'),
		system: buildPptxAiSystemPrompt(),
		tools: await buildPptxAiTools(),
		messages,
	});
	return result.toUIMessageStreamResponse();
}
```

### `model` (en el navegador / local / clave propia) {#model-browser-side-local-bring-your-own-key}

Entregue al visualizador una instancia de modelo que usted construya en el navegador. El asistente ejecuta el bucle de herramientas en el propio proceso, sin backend. Úselo para modelos locales, entornos de desarrollo o entornos de confianza. Tenga en cuenta que la clave está presente en el navegador, por lo que no es adecuado para clientes no confiables.

```ts
import { createOpenAICompatible } from '@ai-sdk/openai-compatible';

const provider = createOpenAICompatible({
	name: 'local',
	baseURL: 'http://localhost:11434/v1', // e.g. Ollama
	apiKey: 'ignored-for-local',
});

const ai: PptxAiConfig = {
	connection: { kind: 'model', model: provider.chatModel('llama3.1') },
};
```

### `transport` (avanzado) {#transport-advanced}

Proporcione un `ChatTransport` del SDK de IA ya construido. Úselo para una pasarela WebSocket, un transporte de flujo de trabajo o pruebas.

```ts
const ai: PptxAiConfig = { connection: { kind: 'transport', transport: myTransport } };
```

## Revisar los cambios: el flujo de proponer y aceptar {#reviewing-changes-the-propose-accept-flow}

De forma predeterminada (`writePolicy: 'stage'`), cada cambio que hace el asistente se **pone en espera**, no se aplica. El panel muestra una tarjeta de propuesta con un resumen breve del cambio y los botones Aceptar, Rechazar y Aceptar todo. Al aceptar, el cambio se confirma mediante el historial de edición del propio visualizador como una única entrada, de modo que un solo Ctrl+Z revierte un cambio de IA aceptado. Al rechazar, se descarta. Las herramientas de lectura (leer diapositivas, buscar texto) y la navegación se ejecutan de inmediato y nunca se ponen en espera.

Opciones de `writePolicy`:

| Valor               | Comportamiento                                                                                                                                        |
| ------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------- |
| `'stage'` (default) | Los cambios se ponen en espera para aceptarlos o rechazarlos de forma explícita.                                                                      |
| `'approve'`         | Usa la solicitud de aprobación por llamada nativa del SDK de IA.                                                                                      |
| `'auto'`            | Los cambios se aplican de inmediato (siguen siendo deshacibles). Las herramientas destructivas, como `delete_slides`, siempre requieren confirmación. |

## Qué puede hacer el asistente (herramientas) {#what-the-assistant-can-do-tools}

Al asistente se le dan un conjunto fijo de herramientas de la presentación. Puede limitarlas con `tools.enabled` (lista de permitidas) o `tools.disabled` (lista de bloqueadas), o añadir las suyas con `tools.extra`.

Hay **55** de ellas, tipadas como `PptxAiToolName`. La mayoría son las mismas funciones que expone el [servidor MCP](/packages/mcp), de modo que un nombre de herramienta significa lo mismo para un asistente integrado en la aplicación que para un agente externo; unas pocas son exclusivas del visualizador porque necesitan la interfaz en vivo (selección, navegación, la vista resuelta de la presentación).

- **Lecturas exclusivas del visualizador** (6): `get_deck_overview`, `get_slide`, `get_element`, `get_speaker_notes`, `find_text`, `get_theme`
- **Navegación y utilidades exclusivas del visualizador** (3): `go_to_slide`, `select_elements`, `merge_tables`
- **Otras lecturas** (7): `get_metadata`, `get_layouts`, `find_placeholders`, `get_presentation_properties`, `run_accessibility_check`, `convert_to_markdown`, `export_to_json`
- **Edición de elementos** (13): `add_element`, `update_element`, `delete_elements`, `arrange_elements`, `clone_element`, `set_element_animation`, `group_elements`, `ungroup_elements`, `batch_update_elements`, `update_element_style`, `replace_geometry`, `set_element_lock`, `manage_hyperlinks`
- **Texto, tablas, gráficos, SmartArt** (11): `replace_text`, `manage_comments`, `update_table_cells`, `manage_table_structure`, `create_chart`, `update_chart`, `add_chart_series`, `remove_chart_series`, `update_chart_series_data`, `manage_smart_art`, `apply_template`
- **Estructura de diapositivas** (6): `add_slide`, `duplicate_slide`, `delete_slides`, `reorder_slides`, `update_slide_properties`, `set_slide_transition`
- **Tema** (3): `apply_theme_preset`, `update_theme_colors`, `update_theme_fonts`
- **Nivel de presentación** (7): `set_canvas_size`, `update_metadata`, `manage_sections`, `manage_custom_shows`, `update_presentation_properties`, `import_from_json`, `apply_layout`

Las notas del orador se pueden leer (`get_speaker_notes`), pero no hay una herramienta de escritura específica para ellas; el asistente edita las notas mediante `update_slide_properties`.

```ts
const ai: PptxAiConfig = {
	connection: { kind: 'endpoint', api: '/api/ai/chat' },
	// Read-only assistant: no editing tools.
	tools: {
		enabled: ['get_deck_overview', 'get_slide', 'get_element', 'find_text', 'go_to_slide'],
	},
};
```

## Contexto de la presentación {#deck-context}

`contextStrategy` controla cuánto de la presentación se envía al modelo en cada turno, además de lo que el modelo obtiene por sí mismo mediante las herramientas de lectura.

| Valor                 | Enviado en cada turno                                                |
| --------------------- | -------------------------------------------------------------------- |
| `'outline'` (default) | Un esquema compacto por diapositiva de toda la presentación.         |
| `'current-slide'`     | Solo la diapositiva activa, en markdown.                             |
| `'none'`              | Nada; el modelo depende por completo de las herramientas de lectura. |

## Otras opciones {#other-options}

```ts
const ai: PptxAiConfig = {
	connection: { kind: 'endpoint', api: '/api/ai/chat' },
	// Appended to the assistant's base instructions (brand voice, guardrails).
	systemPromptExtras: 'Prefer concise edits. Never change the company logo.',
	// Persist and restore a conversation (host-owned storage).
	history: {
		load: (id) => myStore.load(id),
		save: (id, messages) => myStore.save(id, messages),
	},
	onError: (error) => reportToSentry(error),
};
```

## Privacidad {#privacy}

Cuando el asistente se ejecuta, el contenido de la presentación que lee (texto de las diapositivas, estructura y el contexto que permita mediante `contextStrategy`) se envía al proveedor de modelos que haya configurado. Elija el modo de conexión y el proveedor que se ajusten a su política de datos: el modo `endpoint` mantiene su clave y la elección de proveedor en su propia infraestructura, y `tools.enabled` y `contextStrategy: 'none'` le permiten limitar lo que el asistente puede leer.

## Tamaño del paquete {#bundle-size}

Ningún código de IA forma parte de la ruta crítica. El panel de chat y el SDK `ai` se importan de forma dinámica solo cuando el usuario abre el panel por primera vez, y únicamente si pasa una configuración `ai`. Un visualizador sin IA configurada no descarga ningún byte de IA.

## La entrada `/ai` {#ai-entry}

La entrada `/ai` de cada enlace exporta `createAiAssistant(config)`, el tipo de controlador `PptxAiAssistant` y los tipos de configuración `PptxAi*`. El controlador incluye el cargador diferido del panel propio del enlace, el mismo que usa su entrada principal para una configuración simple. Un controlador de Angular incluye solo su configuración: el servicio de chat carga el runtime tanto para un controlador como para una configuración simple.

| Enlace  | Entrada                  | Qué incluye el controlador                                          |
| ------- | ------------------------ | ------------------------------------------------------------------- |
| React   | `pptx-react-viewer/ai`   | el panel de chat con `React.lazy` (`@ai-sdk/react` se carga dentro) |
| Vue     | `pptx-vue-viewer/ai`     | el panel de chat con `defineAsyncComponent` (`@ai-sdk/vue`)         |
| Svelte  | `pptx-svelte-viewer/ai`  | el cargador del componente de panel de chat (`@ai-sdk/svelte`)      |
| Angular | `pptx-angular-viewer/ai` | nada (`AiChatService` importa `ooxml-ui/pptx/ai` por sí mismo)      |
| Vanilla | `pptx-vanilla-viewer/ai` | el cargador del módulo de panel (`ooxml-ui/pptx/ai`)                |

### Lo que garantiza esta versión {#what-this-release-guarantees}

En los cinco enlaces:

- El controlador es un valor opaco y soportado para `ai`. Equivale a pasar su configuración: el mismo panel, el mismo comportamiento. No añade ningún punto de extensión para el host: su tipo no tiene miembros, solo `createAiAssistant` crea uno, y no puede contener un panel ni ningún otro código suyo. Cada enlace guarda sus controladores en un registro privado del paquete, que ninguna exportación expone, de modo que ninguna otra exportación del enlace o de `ooxml-ui` puede crear un controlador que lea el visualizador. La única excepción es interna: la entrada `/ai` de Angular, que ng-packagr construye por su cuenta, accede a ese registro mediante `createAiAssistantHandle`, que el paquete raíz y `/internals` exportan como `@internal`. Solo acepta una configuración y crea el mismo controlador que `createAiAssistant`.
- La entrada `/ai` es estable: `createAiAssistant`, `PptxAiAssistant` y los tipos de configuración `PptxAi*`.
- La configuración simple `PptxAiConfig` es compatible y no está obsoleta. El visualizador lee cualquier valor que no sea un controlador, incluido un objeto con su misma forma, como una configuración simple y la pasa al panel sin validarla, como antes; una configuración que el panel no pueda usar acaba en su estado de error.
- Un controlador creado por otro enlace, o por otra copia del paquete (una segunda instalación, o las compilaciones ESM y CommonJS mezcladas), se lee de la misma forma. En desarrollo, el visualizador avisa de ello una sola vez (`process.env.NODE_ENV !== 'production'`; Angular usa `isDevMode()`).
- Las entradas principales siguen alcanzando el cargador diferido del panel. `pptx-react-viewer` (y `/no-i18next`, las mismas exportaciones), `pptx-vue-viewer` y `pptx-svelte-viewer` con sus entradas `/viewer`, `pptx-angular-viewer` (y `/internals`, el mismo archivo) y `pptx-vanilla-viewer` llegan al paquete SDK del enlace y a `ooxml-ui/pptx/ai` mediante la misma importación dinámica que `/ai`. Ninguno tiene el SDK `ai`, `@ai-sdk/*`, `zod` ni el catálogo de automatización en sus fragmentos iniciales. Por lo tanto, el controlador no cambia qué se carga ni cuándo.

### Por qué el controlador se incluye ahora {#why-the-handle-ships-now}

- Puede cambiar a `createAiAssistant` en una versión menor, de modo que la siguiente versión principal no requiera cambios de código por su parte.
- El controlador mantiene el asistente detrás de una sola importación a la que usted se suscribe. En React, Vue, Svelte y vanilla, el cargador de panel que lleva un controlador procede de la entrada `/ai`, no de la entrada principal. En Angular, el servicio de chat del paquete raíz carga el runtime, de modo que `/ai` nunca entrega código al paquete raíz.

### Lo que cambia en la siguiente versión principal {#what-the-next-major-release-changes}

Las entradas principales dejan de incluir el runtime del asistente cuando se cumplen todas estas condiciones, en cada enlace:

- Los grafos de módulos de código fuente y compilado de las entradas principales (`pptx-react-viewer` y su `/no-i18next`, `pptx-vue-viewer` y `pptx-svelte-viewer` con sus entradas `/viewer`, `pptx-angular-viewer` y su `/internals`, `pptx-vanilla-viewer`) no alcanzan ningún módulo `ai`, `@ai-sdk/*`, `zod`, `ooxml-ui/pptx/ai` ni `ooxml-core/pptx/automation`, ni de forma estática ni mediante una importación dinámica.
- Sus archivos `.d.ts` no alcanzan ningún tipo de `ai` o `@ai-sdk/*`.
- La comprobación de los paquetes de enlace empaquetados (`check-binding-packages.mjs`) espera que solo `./ai`, entre las entradas de cada enlace, alcance el paquete SDK del enlace (`@ai-sdk/react`, `@ai-sdk/vue`, `@ai-sdk/svelte`) y `ooxml-ui/pptx/ai`.
- Una configuración simple en la prop `ai` de una entrada principal se elimina de su tipo y, en tiempo de ejecución, lanza un error que apunta a `createAiAssistant`.

Solo esa versión cambia el comportamiento: `ai` acepta únicamente un controlador, los tipos de configuración `PptxAi*` (que nombran tipos del SDK `ai`) se exportan solo desde las entradas `/ai`, y las entradas principales dejan de tener su propio cargador de panel diferido. El código que ya pasa un controlador seguirá funcionando sin cambios. En Angular, el cargador de runtime está hoy en el paquete raíz, de modo que esa versión también tiene que moverlo a `/ai` sin añadir una exportación que acepte un cargador.
