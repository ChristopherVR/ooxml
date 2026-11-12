# KI-Assistent

Jedes Binding enthält ein optionales, bewusst zu aktivierendes KI-Chatpanel. Ein Benutzer klickt in der Werkzeugleiste auf das KI-Symbol (Funken), und rechts im Viewer öffnet sich ein Chatpanel. Der Assistent kann die Präsentation lesen, Fragen dazu beantworten und Änderungen vorschlagen (Text, Stile, Layout, Folien, Design). Der Benutzer prüft diese Vorschläge und nimmt sie an oder lehnt sie ab, bevor sich etwas ändert.

Das Feature ist **anbieterunabhängig** und **mit eigener Authentifizierung**: Die Bibliothek liefert weder ein Modell noch einen API-Schlüssel mit. Sie verbinden den Assistenten mit einem Modell Ihrer Wahl (einer Backend-Route, einem Modell im Browser oder einem vollständig eigenen Transport). Wenn Sie keine `ai`-Konfiguration übergeben, wird kein Symbol angezeigt und kein KI-Code (und auch nicht das `ai`-SDK) geladen. Ungenutzt verursacht das Feature also keine Kosten.

## Installation {#installation}

Die KI-Pakete sind **optionale Peer-Abhängigkeiten**. Installieren Sie `ai` sowie für React, Vue und Svelte das KI-SDK-UI-Paket des jeweiligen Frameworks. Angular und Vanilla verwenden den im Viewer enthaltenen, framework-unabhängigen Controller und benötigen daher nur `ai`.

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

Außerdem installieren Sie ein **Provider**-Paket für das Modell, das Sie verwenden. Es ist nie eine Abhängigkeit des Viewers, sondern Ihre Wahl:

```bash
# Any OpenAI-compatible endpoint (Ollama, vLLM, LM Studio, a proxy, ...)
bun add @ai-sdk/openai-compatible
# or a first-party provider
bun add @ai-sdk/openai
bun add @ai-sdk/anthropic
```

## Das Panel aktivieren {#enabling-the-panel}

Übergeben Sie einen Assistenten an die Prop `ai` des Viewers (in Angular ein Input, in Vanilla eine Option). Jedes Binding hat einen `/ai`-Eintrag, dessen `createAiAssistant(config)` eine Konfiguration in ein Assistenten-Handle umwandelt. Die Form der Konfiguration ist in allen Bindings gleich; nur die Host-Syntax zum Übergeben einer Prop unterscheidet sich.

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

Die einfache Konfiguration (`ai: { connection }`) funktioniert in jedem Binding weiterhin und ist in dieser Version nicht veraltet. Erzeugen Sie das Handle einmal (auf Modulebene, in `useMemo` oder als Komponentenfeld). Ein bei jedem Render neu erzeugtes Handle hält den geöffneten Chat trotzdem gemountet. Details dazu, was diese Version garantiert und was die nächste Major-Version ändert, finden Sie unter [Der `/ai`-Eintrag](#ai-entry).

## Verbindungsmodi {#connection-modes}

`connection` legt fest, wie der Assistent ein Modell erreicht. Es gibt drei Arten.

### `endpoint` (für den Produktivbetrieb empfohlen) {#endpoint-recommended-for-production}

Senden Sie Nachrichten an eine Route Ihres eigenen Backends. Der Provider-Schlüssel bleibt auf dem Server, der Browser sieht ihn nie. Dies ist der sicherste Standard.

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

Führen Sie auf dem Server das Modell mit dem AI SDK aus und stellen Sie die Präsentations-Werkzeuge als **nur schema-basierte** Werkzeuge bereit. Da diese Werkzeuge kein `execute` besitzen, werden die Werkzeugaufrufe des Modells zum Browser gestreamt, dort gegen die aktive Präsentation ausgeführt und die Ergebnisse anschließend zurück an das Modell gestreamt. Ihr Schlüssel verlässt den Server nie.

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

### `model` (im Browser, lokal, eigener Schlüssel) {#model-browser-side-local-bring-your-own-key}

Übergeben Sie dem Viewer eine Modellinstanz, die Sie im Browser erstellen. Der Assistent führt die Werkzeugschleife im Prozess aus, ohne Backend. Verwenden Sie dies für lokale Modelle, Entwicklungsumgebungen oder vertrauenswürdige Umgebungen. Beachten Sie: Der Schlüssel liegt im Browser, daher ist dies nicht für nicht vertrauenswürdige Clients geeignet.

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

### `transport` (fortgeschritten) {#transport-advanced}

Übergeben Sie einen vollständig konstruierten `ChatTransport` des AI SDK. Verwenden Sie dies für ein WebSocket-Gateway, einen Workflow-Transport oder für Tests.

```ts
const ai: PptxAiConfig = { connection: { kind: 'transport', transport: myTransport } };
```

## Änderungen prüfen: Vorschlag annehmen oder ablehnen {#reviewing-changes-the-propose-accept-flow}

Standardmäßig (`writePolicy: 'stage'`) wird jede Änderung des Assistenten **zur Prüfung vorgemerkt**, statt sofort angewendet zu werden. Das Panel zeigt eine Vorschlagskarte mit einer kurzen Zusammenfassung der Änderung und den Schaltflächen Annehmen, Ablehnen und Alle annehmen. Das Annehmen übernimmt die Änderung über den eigenen Bearbeitungsverlauf des Viewers als einen einzigen Eintrag. Ein einziges Strg+Z macht daher eine angenommene KI-Änderung rückgängig. Ablehnen verwirft die Änderung. Lesende Werkzeuge (Folien lesen, Text suchen) und Navigation laufen sofort und werden nie vorgemerkt.

`writePolicy`-Optionen:

| Wert                 | Verhalten                                                                                                                                                             |
| -------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `'stage'` (Standard) | Schreibvorgänge werden zum ausdrücklichen Annehmen oder Ablehnen vorgemerkt.                                                                                          |
| `'approve'`          | Verwendet die native Freigabe pro Aufruf des AI SDK.                                                                                                                  |
| `'auto'`             | Schreibvorgänge werden sofort angewendet (weiterhin rückgängig machbar). Destruktive Werkzeuge wie `delete_slides` erfordern unabhängig davon immer eine Bestätigung. |

## Was der Assistent kann (Werkzeuge) {#what-the-assistant-can-do-tools}

Der Assistent erhält einen festen Satz an Präsentations-Werkzeugen. Sie können ihn mit `tools.enabled` (Positivliste) bzw. `tools.disabled` (Negativliste) einschränken oder mit `tools.extra` eigene Werkzeuge hinzufügen.

Es gibt **55** davon, typisiert als `PptxAiToolName`. Die meisten sind dieselben Funktionen, die der [MCP-Server](/packages/mcp) bereitstellt. Ein Werkzeugname hat daher für einen eingebetteten Assistenten und für einen externen Agenten dieselbe Bedeutung. Einige wenige sind nur im Viewer verfügbar, weil sie die Live-Oberfläche benötigen (Auswahl, Navigation und die aufgelöste Präsentationsübersicht).

- **Nur im Viewer verfügbare Lesezugriffe** (6): `get_deck_overview`, `get_slide`, `get_element`,
  `get_speaker_notes`, `find_text`, `get_theme`
- **Nur im Viewer verfügbare Navigation und Komfortfunktionen** (3): `go_to_slide`, `select_elements`,
  `merge_tables`
- **Weitere Lesezugriffe** (7): `get_metadata`, `get_layouts`, `find_placeholders`,
  `get_presentation_properties`, `run_accessibility_check`, `convert_to_markdown`,
  `export_to_json`
- **Elementbearbeitung** (13): `add_element`, `update_element`, `delete_elements`,
  `arrange_elements`, `clone_element`, `set_element_animation`, `group_elements`,
  `ungroup_elements`, `batch_update_elements`, `update_element_style`,
  `replace_geometry`, `set_element_lock`, `manage_hyperlinks`
- **Text, Tabellen, Diagramme, SmartArt** (11): `replace_text`, `manage_comments`,
  `update_table_cells`, `manage_table_structure`, `create_chart`, `update_chart`,
  `add_chart_series`, `remove_chart_series`, `update_chart_series_data`,
  `manage_smart_art`, `apply_template`
- **Folienstruktur** (6): `add_slide`, `duplicate_slide`, `delete_slides`,
  `reorder_slides`, `update_slide_properties`, `set_slide_transition`
- **Design** (3): `apply_theme_preset`, `update_theme_colors`, `update_theme_fonts`
- **Präsentationsebene** (7): `set_canvas_size`, `update_metadata`, `manage_sections`, `manage_custom_shows`,
  `update_presentation_properties`, `import_from_json`, `apply_layout`

Sprechernotizen sind lesbar (`get_speaker_notes`), aber es gibt dafür kein eigenes Schreibwerkzeug. Der Assistent bearbeitet Notizen über `update_slide_properties`.

```ts
const ai: PptxAiConfig = {
	connection: { kind: 'endpoint', api: '/api/ai/chat' },
	// Read-only assistant: no editing tools.
	tools: {
		enabled: ['get_deck_overview', 'get_slide', 'get_element', 'find_text', 'go_to_slide'],
	},
};
```

## Präsentationskontext {#deck-context}

`contextStrategy` steuert, wie viel von der Präsentation pro Runde zusätzlich zu dem an das Modell gesendet wird, was es selbst über die Lesewerkzeuge abruft.

| Wert                   | Pro Runde gesendet                                                  |
| ---------------------- | ------------------------------------------------------------------- |
| `'outline'` (Standard) | Eine kompakte Gliederung der gesamten Präsentation, folienweise.    |
| `'current-slide'`      | Nur die aktive Folie, als Markdown.                                 |
| `'none'`               | Nichts; das Modell verlässt sich vollständig auf die Lesewerkzeuge. |

## Weitere Optionen {#other-options}

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

## Datenschutz {#privacy}

Wenn der Assistent läuft, wird der Präsentationsinhalt, den er liest (Folientext, Struktur und der Kontext, den Sie über `contextStrategy` erlauben), an den Modellanbieter gesendet, den Sie konfiguriert haben. Wählen Sie Verbindungsmodus und Anbieter passend zu Ihrer Datenrichtlinie. Der Modus `endpoint` hält Schlüssel und Anbieterwahl auf Ihrer eigenen Infrastruktur. Mit `tools.enabled` und `contextStrategy: 'none'` begrenzen Sie, was der Assistent lesen kann.

## Bundle-Größe {#bundle-size}

Kein KI-Code gehört zum kritischen Ladepfad. Das Chatpanel und das `ai`-SDK werden erst dynamisch importiert, wenn der Benutzer das Panel zum ersten Mal öffnet, und nur dann, wenn Sie überhaupt eine `ai`-Konfiguration übergeben. Ein Viewer ohne konfigurierte KI lädt keinerlei KI-Bytes.

## Der `/ai`-Eintrag {#ai-entry}

Der `/ai`-Eintrag jedes Bindings exportiert `createAiAssistant(config)`, den Handle-Typ `PptxAiAssistant` und die Konfigurationstypen `PptxAi*`. Das Handle enthält den eigenen Lazy-Loader für das Panel des Bindings, also denselben, den der Haupteintrag auch für eine einfache Konfiguration verwendet. Ein Angular-Handle enthält nur seine Konfiguration. Der Chat-Dienst lädt die Laufzeit sowohl für ein Handle als auch für eine einfache Konfiguration.

| Binding | Eintrag                  | Was das Handle enthält                                          |
| ------- | ------------------------ | --------------------------------------------------------------- |
| React   | `pptx-react-viewer/ai`   | das `React.lazy`-Chatpanel (`@ai-sdk/react` wird darin geladen) |
| Vue     | `pptx-vue-viewer/ai`     | das `defineAsyncComponent`-Chatpanel (`@ai-sdk/vue`)            |
| Svelte  | `pptx-svelte-viewer/ai`  | den Loader der Chatpanel-Komponente (`@ai-sdk/svelte`)          |
| Angular | `pptx-angular-viewer/ai` | nichts (`AiChatService` importiert `ooxml-ui/pptx/ai` selbst)   |
| Vanilla | `pptx-vanilla-viewer/ai` | den Loader des Panelmoduls (`ooxml-ui/pptx/ai`)                 |

### Was diese Version garantiert {#what-this-release-guarantees}

In allen fünf Bindings gilt:

- Das Handle ist ein undurchsichtiger, unterstützter Wert für `ai`. Es entspricht der Übergabe seiner Konfiguration: dasselbe Panel, dasselbe Verhalten. Es bietet keinen Erweiterungspunkt für den Host. Sein Typ hat keine Mitglieder, nur `createAiAssistant` erzeugt eines, und es kann kein eigenes Panel oder anderen eigenen Code enthalten. Jedes Binding hält seine Handles in einer Registry, die nur dem Paket selbst gehört und die kein Eintrag exportiert. Deshalb kann kein anderer Export des Bindings oder von `ooxml-ui` ein Handle erzeugen, das der Viewer liest. Die einzige Ausnahme ist intern: Der `/ai`-Eintrag von Angular, den ng-packagr selbst baut, erreicht diese Registry über `createAiAssistantHandle`. Dieses ist im Paket-Root und unter `/internals` als `@internal` exportiert. Es nimmt nur eine Konfiguration entgegen und erzeugt dasselbe Handle wie `createAiAssistant`.
- Der `/ai`-Eintrag ist stabil: `createAiAssistant`, `PptxAiAssistant` und die `PptxAi*`-Konfigurationstypen.
- Die einfache `PptxAiConfig` wird unterstützt und ist nicht veraltet. Der Viewer liest jeden Wert, der kein Handle ist, auch ein Objekt mit derselben Form, als einfache Konfiguration. Er übergibt ihn wie bisher ungeprüft an das Panel. Eine Konfiguration, die das Panel nicht verwenden kann, endet in dessen Fehlerzustand.
- Ein Handle, das von einem anderen Binding oder von einer anderen Kopie des Pakets stammt (eine zweite Installation oder gemischte ESM- und CommonJS-Builds), wird genauso gelesen. In der Entwicklung warnt der Viewer einmalig davor (`process.env.NODE_ENV !== 'production'`; Angular verwendet `isDevMode()`).
- Die Haupteinstiege erreichen weiterhin den Lazy-Loader des Panels. `pptx-react-viewer` (und `/no-i18next` mit denselben Exporten), `pptx-vue-viewer` und `pptx-svelte-viewer` mit ihren `/viewer`-Einstiegen, `pptx-angular-viewer` (und `/internals`, dieselbe Datei) sowie `pptx-vanilla-viewer` erreichen das SDK-Paket des Bindings und `ooxml-ui/pptx/ai` über denselben dynamischen Import wie `/ai`. Keiner von ihnen enthält das `ai`-SDK, `@ai-sdk/*`, `zod` oder den Automatisierungskatalog in seinen initialen Chunks. Das Handle ändert daher nicht, was geladen wird oder wann.

### Warum das Handle jetzt ausgeliefert wird {#why-the-handle-ships-now}

- Sie können in einer Minor-Version auf `createAiAssistant` umstellen, sodass die nächste Major-Version auf Ihrer Seite keine Codeänderung erfordert.
- Das Handle hält den Assistenten hinter einem Import, den Sie bewusst wählen. In React, Vue, Svelte und Vanilla stammt der Panel-Loader, den ein Handle enthält, aus dem `/ai`-Eintrag und nicht aus dem Haupteintrag. In Angular lädt der Chat-Dienst im Paket-Root die Laufzeit, daher gibt `/ai` dem Root nie Code weiter.

### Was die nächste Major-Version ändert {#what-the-next-major-release-changes}

Die Haupteinstiege sind in jedem Binding frei von der Assistenten-Laufzeit, wenn alle folgenden Punkte zutreffen:

- Die Quell- und Build-Modulgraphen der Haupteinstiege (`pptx-react-viewer` und `/no-i18next`, `pptx-vue-viewer` und `pptx-svelte-viewer` mit ihren `/viewer`-Einstiegen, `pptx-angular-viewer` und `/internals`, `pptx-vanilla-viewer`) erreichen kein Modul `ai`, `@ai-sdk/*`, `zod`, `ooxml-ui/pptx/ai` oder `ooxml-core/pptx/automation`, weder statisch noch über einen dynamischen Import.
- Ihre `.d.ts`-Dateien erreichen keinen Typ aus `ai` oder `@ai-sdk/*`.
- Die Prüfung der gepackten Bindings (`check-binding-packages.mjs`) erwartet nur für `./ai` unter den Einstiegen jedes Bindings, dass es das SDK-Paket des Bindings (`@ai-sdk/react`, `@ai-sdk/vue`, `@ai-sdk/svelte`) und `ooxml-ui/pptx/ai` erreicht.
- Eine einfache Konfiguration am `ai` eines Haupteinstiegs wird aus dessen Typ entfernt und wirft zur Laufzeit einen Fehler, der auf `createAiAssistant` verweist.

Nur diese Version ändert das Verhalten: `ai` akzeptiert dann nur noch ein Handle, die `PptxAi*`-Konfigurationstypen (die Typen des `ai`-SDK benennen) werden nur noch aus den `/ai`-Einstiegen exportiert, und die Haupteinstiege verzichten auf ihren eigenen Lazy-Loader für das Panel. Code, der bereits ein Handle übergibt, funktioniert unverändert weiter. In Angular liegt der Laufzeit-Loader heute im Paket-Root. Daher muss diese Version ihn auch nach `/ai` verlegen, ohne einen Export hinzuzufügen, der einen Loader entgegennimmt.
