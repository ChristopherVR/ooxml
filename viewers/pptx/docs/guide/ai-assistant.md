# AI Assistant

Every binding ships an optional, opt-in AI chat panel. A user clicks the AI
(sparkles) icon in the toolbar and a chat panel opens on the right side of the
viewer. The assistant can read the deck, answer questions about it, and propose
edits (text, styles, layout, slides, theme) that the user reviews and accepts or
rejects before anything changes.

The feature is **provider-agnostic** and **bring-your-own-auth**: the library
ships no model and no API key. You wire the assistant to a model you control (a
backend route, a browser-side model, or a fully custom transport). If you do not
pass an `ai` config, no icon renders and none of the AI code (or the `ai` SDK)
is loaded, so there is zero cost when the feature is unused.

## Installation

The AI packages are **optional peer dependencies**. Install `ai` plus, for
React / Vue / Svelte, that framework's AI SDK UI package. Angular and Vanilla
use the framework-agnostic controller bundled in the viewer, so they only need
`ai`.

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

You also install one **provider** package for whichever model you use. This is
never a dependency of the viewer; it is your choice:

```bash
# Any OpenAI-compatible endpoint (Ollama, vLLM, LM Studio, a proxy, ...)
bun add @ai-sdk/openai-compatible
# or a first-party provider
bun add @ai-sdk/openai
bun add @ai-sdk/anthropic
```

## Enabling the panel

Pass an assistant to the viewer's `ai` prop (an input in Angular, an option in vanilla).
Every binding has an `/ai` entry whose `createAiAssistant(config)` turns a config into an
assistant handle. The config shape is identical across every binding; only the host
syntax for passing a prop differs.

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

Passing the plain config instead (`ai: { connection }`) keeps working in every binding and
is not deprecated in this release. Create the handle once (at module scope, in `useMemo`,
or as a component field); a handle recreated on every render still keeps the open chat
mounted. See [The `/ai` entry](#ai-entry) for what this release guarantees and what the
next major release changes.

## Connection modes

`connection` declares how the assistant reaches a model. There are three kinds.

### `endpoint` (recommended for production)

Post messages to a route on your own backend. The provider key stays
server-side; the browser never sees it. This is the safest default.

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

On the server, run the model with the AI SDK and expose the deck tools as
**schema-only** tools. Because the tools carry no `execute`, the model's tool
calls are streamed back to the browser and run against the live deck there, then
the results stream back to the model. Your key never leaves the server.

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

### `model` (browser-side / local / bring-your-own key)

Hand the viewer a model instance you construct in the browser. The assistant
runs the tool loop in-process, with no backend. Use this for local models, dev
setups, or trusted environments. Note the key is present in the browser, so this
is not appropriate for untrusted clients.

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

### `transport` (advanced)

Provide a fully-constructed AI SDK `ChatTransport`. Use this for a WebSocket
gateway, a workflow transport, or tests.

```ts
const ai: PptxAiConfig = { connection: { kind: 'transport', transport: myTransport } };
```

## Reviewing changes: the propose / accept flow

By default (`writePolicy: 'stage'`) every edit the assistant makes is **staged**,
not applied. The panel shows a proposal card with a short summary of the change
and Apply / Discard buttons, plus Apply all when more than one change can be
applied. Applying a slide change records one undo step, so one Ctrl+Z reverts it,
and Apply all records one step per change. A theme change re-colours the slides at
once, is saved with the file and is one undo step in every viewer. Other
presentation-level changes (document properties, sections, slide size) behave like
the same edit made by hand in that viewer, so whether Ctrl+Z reverts a document
property change depends on the viewer. Read tools (reading slides, finding text)
and navigation run immediately and are never staged.

Proposals belong to the presentation they were made on. Editing the deck keeps
them, Slide Master view commands and slide paste included, and they still apply
afterwards. Loading a presentation discards them: opening a file (the same file
again too), starting a new presentation, restoring a recovered version, or
joining a collaboration session whose room replaces the slides. The panel then
says the suggested changes were discarded.
Nothing the assistant wrote or staged on one presentation reaches the next.

`writePolicy` applies to every tool that changes the deck, theme and
presentation-level tools included:

| Value               | Behavior                                                                   |
| ------------------- | -------------------------------------------------------------------------- |
| `'stage'` (default) | Writes wait for review. The user applies each one, or all of them at once. |
| `'approve'`         | Writes wait for review, and each one must be applied on its own.           |
| `'auto'`            | Writes apply immediately (still undoable).                                 |

`delete_slides` and `import_from_json` always wait until the user applies that one
change, under every policy. Apply all leaves them pending.

A proposal records what it changes, not a copy of the deck. Applying it replays
the change on the deck as it is at that moment, so several proposals and anything
the user edited in the meantime all survive. When the same content has changed
since the proposal was made, it is not applied: the card reports the conflict and
the user can discard it or try again.

The viewers' bridges let each applied change read what the previous one wrote,
even before the framework re-renders. If you implement `PptxAiBridge` yourself,
`checkPptxAiBridgeReadsItsWrites(bridge)` from `ooxml-ui/pptx/ai/host` reports any
read that misses a write. Run it on a throwaway deck, because it edits the deck.
A proposal passes `applyTheme` only the theme parts it changed. A bridge's
`getDeckLoad()` names the presentation load: return a new value whenever your host
loads a presentation, and the same value through every edit, even one that
replaces the handler.

The assistant's tool arguments are checked against each tool's schema before the
tool runs. A call that does not match is returned to the model as an error and
changes nothing.

## Run limits {#run-limits}

One user message starts a run, which lasts until the model stops calling tools. A
run ends early, with a message in the panel, when it reaches a limit:

```ts
const ai: PptxAiConfig = {
	connection: { kind: 'endpoint', api: '/api/ai/chat' },
	limits: {
		maxSteps: 16, // model requests
		maxToolCalls: 48,
		maxRunMs: 600_000,
		maxRepeatedToolCalls: 4, // the same tool with the same input
	},
};
```

The values above are the defaults. The panel also says so when a reply was cut off
by the length limit, stopped by a content filter or ended without a finish signal.
`session.getLastFinish()` (and `getLastFinish()` on the Vanilla chat controller)
returns how the latest request ended: `reason` is one of `stop`, `tool-calls`,
`length`, `filtered`, `cancelled`, `error`, `limit` or `unknown`, with `limit` naming
`steps`, `tool-calls`, `time` or `repeats` when one ended the run.

## What the assistant can do (tools)

The assistant is given a fixed set of deck tools. You can narrow them with
`tools.enabled` (allowlist) / `tools.disabled` (denylist), or add your own with
`tools.extra`.

There are **55** of them, typed as `PptxAiToolName`. Most are the same functions the
[MCP server](/packages/mcp) exposes, so a tool name means the same thing to an in-app
assistant and to an external agent; a handful are viewer-only because they need the live
UI (selection, navigation, the resolved deck overview).

- **Viewer-only reads** (6): `get_deck_overview`, `get_slide`, `get_element`,
  `get_speaker_notes`, `find_text`, `get_theme`
- **Viewer-only navigation and convenience** (3): `go_to_slide`, `select_elements`,
  `merge_tables`
- **Other reads** (7): `get_metadata`, `get_layouts`, `find_placeholders`,
  `get_presentation_properties`, `run_accessibility_check`, `convert_to_markdown`,
  `export_to_json`
- **Element editing** (13): `add_element`, `update_element`, `delete_elements`,
  `arrange_elements`, `clone_element`, `set_element_animation`, `group_elements`,
  `ungroup_elements`, `batch_update_elements`, `update_element_style`,
  `replace_geometry`, `set_element_lock`, `manage_hyperlinks`
- **Text, tables, charts, SmartArt** (11): `replace_text`, `manage_comments`,
  `update_table_cells`, `manage_table_structure`, `create_chart`, `update_chart`,
  `add_chart_series`, `remove_chart_series`, `update_chart_series_data`,
  `manage_smart_art`, `apply_template`
- **Slide structure** (6): `add_slide`, `duplicate_slide`, `delete_slides`,
  `reorder_slides`, `update_slide_properties`, `set_slide_transition`
- **Theme** (3, reviewed like every other edit): `apply_theme_preset`, `update_theme_colors`, `update_theme_fonts`
- **Presentation level** (7): `set_canvas_size`, `update_metadata`, `manage_sections`, `manage_custom_shows`,
  `update_presentation_properties`, `import_from_json`, `apply_layout`

Speaker notes are readable (`get_speaker_notes`) but there is no dedicated write tool for
them; the assistant edits notes through `update_slide_properties`.

```ts
const ai: PptxAiConfig = {
	connection: { kind: 'endpoint', api: '/api/ai/chat' },
	// Read-only assistant: no editing tools.
	tools: {
		enabled: ['get_deck_overview', 'get_slide', 'get_element', 'find_text', 'go_to_slide'],
	},
};
```

## Deck context

`contextStrategy` controls how much of the deck is sent to the model with each
request in every binding, in addition to what the model fetches itself via the read
tools. What the user has focused on is sent as well, whatever the strategy.

| Value                 | Sent each turn                                    |
| --------------------- | ------------------------------------------------- |
| `'outline'` (default) | A compact per-slide outline of the whole deck.    |
| `'current-slide'`     | Only the active slide, as markdown.               |
| `'none'`              | Nothing; the model relies entirely on read tools. |

## Other options

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

## Privacy

When the assistant runs, the deck content it reads (slide text, structure, and
the context you allow via `contextStrategy`) is sent to whichever model provider
you configured. Choose the connection mode and provider that match your data
policy: the `endpoint` mode keeps your key and provider choice on your own
infrastructure, and `tools.enabled` / `contextStrategy: 'none'` let you limit
what the assistant can read.

## Bundle size

None of the AI code ships on the critical path. The chat panel and the `ai` SDK
are dynamically imported only when the user first opens the panel, and only when
you pass an `ai` config at all. A viewer without AI configured downloads zero AI
bytes.

## The `/ai` entry {#ai-entry}

Each binding's `/ai` entry exports `createAiAssistant(config)`, the `PptxAiAssistant`
handle type and the `PptxAi*` config types. The handle carries the binding's own lazy
panel loader, the same one its main entry uses for a plain config. An Angular handle
carries its config only: the chat service loads the runtime for a handle and a plain
config alike.

| Binding | Entry                    | What the handle carries                                       |
| ------- | ------------------------ | ------------------------------------------------------------- |
| React   | `pptx-react-viewer/ai`   | the `React.lazy` chat panel (`@ai-sdk/react` loads inside it) |
| Vue     | `pptx-vue-viewer/ai`     | the `defineAsyncComponent` chat panel (`@ai-sdk/vue`)         |
| Svelte  | `pptx-svelte-viewer/ai`  | the loader of the chat panel component (`@ai-sdk/svelte`)     |
| Angular | `pptx-angular-viewer/ai` | nothing (`AiChatService` imports `ooxml-ui/pptx/ai` itself)   |
| Vanilla | `pptx-vanilla-viewer/ai` | the loader of the panel module (`ooxml-ui/pptx/ai`)           |

### What this release guarantees {#what-this-release-guarantees}

In all five bindings:

- The handle is an opaque, supported value for `ai`. It is equivalent to passing its
  config: the same panel, the same behaviour. It adds no host extension point: its type
  has no members, only `createAiAssistant` makes one, and it cannot carry a panel or any
  other code of yours. Each binding keeps its handles in a registry private to the package,
  which no entry exports, so no other export of the binding or of `ooxml-ui` can make a
  handle the viewer reads. The one exception is internal: Angular's `/ai` entry, which
  ng-packagr builds on its own, reaches that registry through `createAiAssistantHandle`,
  which the package root and `/internals` export as `@internal`. It takes a config only and
  makes the same handle as `createAiAssistant`.
- The `/ai` entry is stable: `createAiAssistant`, `PptxAiAssistant` and the `PptxAi*`
  config types.
- The plain `PptxAiConfig` is supported and not deprecated. The viewer reads any value
  that is not a handle, including an object shaped like one, as a plain config and passes
  it to the panel without validating it, as before; a config the panel cannot use ends in
  its error state.
- A handle made by another binding, or by another copy of the package (a second install,
  or the ESM and CommonJS builds mixed), is read the same way. In development the viewer
  warns about it once (`process.env.NODE_ENV !== 'production'`; Angular uses
  `isDevMode()`).
- The main entries still reach the lazy panel loader. `pptx-react-viewer` (and
  `/no-i18next`, the same exports), `pptx-vue-viewer` and `pptx-svelte-viewer` with their
  `/viewer` entries, `pptx-angular-viewer` (and `/internals`, the same file) and `pptx-vanilla-viewer` reach
  the binding's SDK package and `ooxml-ui/pptx/ai` through the same dynamic import as
  `/ai`. None of them has the `ai` SDK, `@ai-sdk/*`, `zod` or the automation catalogue in
  its initial chunks. The handle therefore does not change what loads, or when.

### Why the handle ships now {#why-the-handle-ships-now}

- You can switch to `createAiAssistant` in a minor release, so the next major release
  needs no code change on your side.
- The handle keeps the assistant behind one import you opt into. In React, Vue, Svelte and
  vanilla, the panel loader a handle carries comes from the `/ai` entry, not from the main
  entry. In Angular the chat service in the package root loads the runtime, so that `/ai`
  never hands the root any code.

### What the next major release changes {#what-the-next-major-release-changes}

The main entries become free of the assistant runtime when all of these hold, in every
binding:

- The source and built module graphs of the main entries (`pptx-react-viewer` and its
  `/no-i18next`, `pptx-vue-viewer` and `pptx-svelte-viewer` with their `/viewer` entries,
  `pptx-angular-viewer` and its `/internals`, `pptx-vanilla-viewer`) reach no `ai`,
  `@ai-sdk/*`, `zod`, `ooxml-ui/pptx/ai` or `ooxml-core/pptx/automation` module,
  statically or through a dynamic import.
- Their `.d.ts` files reach no type from `ai` or `@ai-sdk/*`.
- The packed-binding check (`check-binding-packages.mjs`) expects only `./ai` among each
  binding's entries to reach the binding's SDK package (`@ai-sdk/react`, `@ai-sdk/vue`,
  `@ai-sdk/svelte`) and `ooxml-ui/pptx/ai`.
- A plain config on a main entry's `ai` is removed from its type, and at runtime throws an
  error that points to `createAiAssistant`.

Only that release changes behaviour: `ai` takes only a handle, the `PptxAi*` config types
(which name `ai` SDK types) are exported only from the `/ai` entries, and the main entries
drop their own lazy panel loader. Code that already passes a handle keeps working
unchanged. In Angular the runtime loader sits in the package root today, so that release
also has to move it into `/ai` without adding an export that takes a loader.
