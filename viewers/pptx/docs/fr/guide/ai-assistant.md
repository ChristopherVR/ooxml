# Assistant IA

Chaque binding fournit un panneau de discussion IA facultatif, activé explicitement. L'utilisateur clique sur l'icône IA (étincelles) de la barre d'outils et un panneau de discussion s'ouvre sur le côté droit de la visionneuse. L'assistant peut lire la présentation, répondre à des questions à son sujet et proposer des modifications (texte, styles, mise en page, diapositives, thème) que l'utilisateur examine, puis accepte ou rejette avant que quoi que ce soit ne change.

La fonctionnalité est **indépendante du fournisseur** et **avec authentification à votre charge** : la bibliothèque ne fournit ni modèle ni clé d'API. Vous branchez l'assistant sur un modèle que vous contrôlez (une route de backend, un modèle côté navigateur ou un transport entièrement personnalisé). Si vous ne passez pas de configuration `ai`, aucune icône ne s'affiche et aucun code IA (ni le SDK `ai`) n'est chargé : la fonctionnalité n'a aucun coût lorsqu'elle n'est pas utilisée.

## Installation {#installation}

Les packages IA sont des **dépendances peer optionnelles**. Installez `ai` ainsi que, pour React / Vue / Svelte, le package UI du SDK IA de ce framework. Angular et Vanilla utilisent le contrôleur indépendant du framework inclus dans la visionneuse : ils n'ont besoin que de `ai`.

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

Vous installez aussi un **fournisseur** pour le modèle que vous utilisez. Ce n'est jamais une dépendance de la visionneuse ; le choix vous appartient :

```bash
# Any OpenAI-compatible endpoint (Ollama, vLLM, LM Studio, a proxy, ...)
bun add @ai-sdk/openai-compatible
# or a first-party provider
bun add @ai-sdk/openai
bun add @ai-sdk/anthropic
```

## Activer le panneau {#enabling-the-panel}

Passez un assistant à la prop `ai` de la visionneuse (une entrée dans Angular, une option en vanilla). Chaque binding possède une entrée `/ai` dont `createAiAssistant(config)` transforme une configuration en un handle d'assistant. La forme de la configuration est identique dans chaque binding ; seule la syntaxe de l'hôte pour passer une prop change.

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

Passer la configuration simple à la place (`ai: { connection }`) continue de fonctionner dans chaque binding et n'est pas déprécié dans cette version. Créez le handle une seule fois (au niveau du module, dans `useMemo` ou comme champ de composant) : un handle recréé à chaque rendu empêche le maintien du chat ouvert. Voir [L'entrée `/ai`](#ai-entry) pour ce que cette version garantit et ce que la prochaine version majeure change.

## Modes de connexion {#connection-modes}

`connection` déclare comment l'assistant atteint un modèle. Il existe trois types.

### `endpoint` (recommandé en production) {#endpoint-recommended-for-production}

Envoyez les messages à une route de votre propre backend. La clé du fournisseur reste côté serveur ; le navigateur ne la voit jamais. C'est la valeur par défaut la plus sûre.

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

Côté serveur, exécutez le modèle avec le SDK IA et exposez les outils de présentation sous forme d'outils **à schéma seul**. Comme ces outils n'ont pas de `execute`, les appels d'outils du modèle sont renvoyés en flux vers le navigateur, exécutés sur la présentation active à cet endroit, puis les résultats repartent en flux vers le modèle. Votre clé ne quitte jamais le serveur.

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

### `model` (côté navigateur / local / clé fournie par vous) {#model-browser-side-local-bring-your-own-key}

Donnez à la visionneuse une instance de modèle construite dans le navigateur. L'assistant exécute la boucle d'outils dans le processus, sans backend. Utilisez ce mode pour des modèles locaux, des environnements de développement ou des environnements de confiance. Attention : la clé est présente dans le navigateur, ce mode ne convient donc pas aux clients non fiables.

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

### `transport` (avancé) {#transport-advanced}

Fournissez un `ChatTransport` du SDK IA entièrement construit. Utilisez-le pour une passerelle WebSocket, un transport de workflow ou des tests.

```ts
const ai: PptxAiConfig = { connection: { kind: 'transport', transport: myTransport } };
```

## Examiner les modifications : le flux proposer / accepter {#reviewing-changes-the-propose-accept-flow}

Par défaut (`writePolicy: 'stage'`), chaque modification de l'assistant est **mise en attente**, et non appliquée. Le panneau affiche une carte de proposition avec un bref résumé de la modification et les boutons Accepter / Rejeter / Tout accepter. Accepter valide la modification via l'historique d'édition propre à la visionneuse, en une seule entrée : un seul Ctrl+Z annule une modification IA acceptée. Rejeter l'abandonne. Les outils de lecture (lecture des diapositives, recherche de texte) et la navigation s'exécutent immédiatement et ne sont jamais mis en attente.

Options de `writePolicy` :

| Valeur             | Comportement                                                                                                                                    |
| ------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| `'stage'` (défaut) | Les écritures sont mises en attente d'une acceptation ou d'un rejet explicite.                                                                  |
| `'approve'`        | Utilise la demande d'approbation native du SDK IA, appel par appel.                                                                             |
| `'auto'`           | Les écritures s'appliquent immédiatement (restent annulables). Les outils destructeurs comme `delete_slides` exigent toujours une confirmation. |

## Ce que l'assistant peut faire (outils) {#what-the-assistant-can-do-tools}

L'assistant reçoit un ensemble fixe d'outils de présentation. Vous pouvez le restreindre avec `tools.enabled` (liste d'autorisation) / `tools.disabled` (liste de refus), ou ajouter les vôtres avec `tools.extra`.

Il existe **55** outils, typés `PptxAiToolName`. La plupart sont les mêmes fonctions que celles exposées par le [serveur MCP](/packages/mcp) : un nom d'outil a donc le même sens pour un assistant intégré à l'application et pour un agent externe. Quelques-uns sont propres à la visionneuse, car ils ont besoin de l'interface en direct (sélection, navigation, vue d'ensemble résolue de la présentation).

- **Lectures propres à la visionneuse** (6) : `get_deck_overview`, `get_slide`, `get_element`,
  `get_speaker_notes`, `find_text`, `get_theme`
- **Navigation et commodités propres à la visionneuse** (3) : `go_to_slide`, `select_elements`,
  `merge_tables`
- **Autres lectures** (7) : `get_metadata`, `get_layouts`, `find_placeholders`,
  `get_presentation_properties`, `run_accessibility_check`, `convert_to_markdown`,
  `export_to_json`
- **Modification d'éléments** (13) : `add_element`, `update_element`, `delete_elements`,
  `arrange_elements`, `clone_element`, `set_element_animation`, `group_elements`,
  `ungroup_elements`, `batch_update_elements`, `update_element_style`,
  `replace_geometry`, `set_element_lock`, `manage_hyperlinks`
- **Texte, tableaux, graphiques, SmartArt** (11) : `replace_text`, `manage_comments`,
  `update_table_cells`, `manage_table_structure`, `create_chart`, `update_chart`,
  `add_chart_series`, `remove_chart_series`, `update_chart_series_data`,
  `manage_smart_art`, `apply_template`
- **Structure des diapositives** (6) : `add_slide`, `duplicate_slide`, `delete_slides`,
  `reorder_slides`, `update_slide_properties`, `set_slide_transition`
- **Thème** (3) : `apply_theme_preset`, `update_theme_colors`, `update_theme_fonts`
- **Niveau présentation** (7) : `set_canvas_size`, `update_metadata`, `manage_sections`, `manage_custom_shows`,
  `update_presentation_properties`, `import_from_json`, `apply_layout`

Les notes de l'orateur sont lisibles (`get_speaker_notes`), mais aucun outil d'écriture dédié n'existe pour elles : l'assistant modifie les notes via `update_slide_properties`.

```ts
const ai: PptxAiConfig = {
	connection: { kind: 'endpoint', api: '/api/ai/chat' },
	// Read-only assistant: no editing tools.
	tools: {
		enabled: ['get_deck_overview', 'get_slide', 'get_element', 'find_text', 'go_to_slide'],
	},
};
```

## Contexte de la présentation {#deck-context}

`contextStrategy` contrôle la quantité de présentation envoyée au modèle à chaque tour, en plus de ce que le modèle récupère lui-même via les outils de lecture.

| Valeur               | Envoyé à chaque tour                                                    |
| -------------------- | ----------------------------------------------------------------------- |
| `'outline'` (défaut) | Un plan compact, diapositive par diapositive, de toute la présentation. |
| `'current-slide'`    | Seulement la diapositive active, en markdown.                           |
| `'none'`             | Rien ; le modèle s'appuie entièrement sur les outils de lecture.        |

## Autres options {#other-options}

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

## Confidentialité {#privacy}

Lorsque l'assistant s'exécute, le contenu de présentation qu'il lit (texte des diapositives, structure, et le contexte que vous autorisez via `contextStrategy`) est envoyé au fournisseur de modèle que vous avez configuré. Choisissez le mode de connexion et le fournisseur qui correspondent à votre politique de données : le mode `endpoint` garde votre clé et le choix du fournisseur sur votre propre infrastructure, et `tools.enabled` / `contextStrategy: 'none'` vous permettent de limiter ce que l'assistant peut lire.

## Taille du bundle {#bundle-size}

Aucun code IA ne fait partie du chemin critique. Le panneau de discussion et le SDK `ai` ne sont importés dynamiquement que lorsque l'utilisateur ouvre le panneau pour la première fois, et seulement si vous passez une configuration `ai`. Une visionneuse sans IA configurée ne télécharge aucun octet IA.

## L'entrée `/ai` {#ai-entry}

L'entrée `/ai` de chaque binding exporte `createAiAssistant(config)`, le type de handle `PptxAiAssistant` et les types de configuration `PptxAi*`. Le handle transporte le chargeur paresseux de panneau propre au binding, le même que celui que son entrée principale utilise pour une configuration simple. Un handle Angular ne transporte que sa configuration : le service de chat charge le runtime, qu'il reçoive un handle ou une configuration simple.

| Binding | Entrée                   | Ce que transporte le handle                                                  |
| ------- | ------------------------ | ---------------------------------------------------------------------------- |
| React   | `pptx-react-viewer/ai`   | le panneau de chat en `React.lazy` (`@ai-sdk/react` se charge à l'intérieur) |
| Vue     | `pptx-vue-viewer/ai`     | le panneau de chat en `defineAsyncComponent` (`@ai-sdk/vue`)                 |
| Svelte  | `pptx-svelte-viewer/ai`  | le chargeur du composant de panneau de chat (`@ai-sdk/svelte`)               |
| Angular | `pptx-angular-viewer/ai` | rien (`AiChatService` importe lui-même `ooxml-ui/pptx/ai`)                   |
| Vanilla | `pptx-vanilla-viewer/ai` | le chargeur du module de panneau (`ooxml-ui/pptx/ai`)                        |

### Ce que cette version garantit {#what-this-release-guarantees}

Dans les cinq bindings :

- Le handle est une valeur opaque et prise en charge pour `ai`. Il équivaut à passer sa configuration : même panneau, même comportement. Il n'ajoute aucun point d'extension pour l'hôte : son type n'a aucun membre, seul `createAiAssistant` en crée un, et il ne peut transporter ni panneau ni aucun autre code à vous. Chaque binding conserve ses handles dans un registre privé au package, qu'aucune entrée n'exporte ; aucun autre export du binding ou de `ooxml-ui` ne peut donc fabriquer un handle que la visionneuse lit. L'exception est interne : l'entrée `/ai` d'Angular, que ng-packagr construit seul, atteint ce registre via `createAiAssistantHandle`, que la racine du package et `/internals` exportent comme `@internal`. Elle ne prend qu'une configuration et produit le même handle que `createAiAssistant`.
- L'entrée `/ai` est stable : `createAiAssistant`, `PptxAiAssistant` et les types de configuration `PptxAi*`.
- La configuration `PptxAiConfig` simple est prise en charge et n'est pas dépréciée. La visionneuse lit toute valeur qui n'est pas un handle, y compris un objet qui en a la forme, comme une configuration simple et la transmet au panneau sans la valider, comme auparavant ; une configuration que le panneau ne peut pas utiliser aboutit à son état d'erreur.
- Un handle produit par un autre binding, ou par une autre copie du package (une seconde installation, ou des builds ESM et CommonJS mélangés), est lu de la même façon. En développement, la visionneuse avertit une fois à ce sujet (`process.env.NODE_ENV !== 'production'` ; Angular utilise `isDevMode()`).
- Les entrées principales atteignent toujours le chargeur paresseux de panneau. `pptx-react-viewer` (et `/no-i18next`, mêmes exports), `pptx-vue-viewer` et `pptx-svelte-viewer` avec leurs entrées `/viewer`, `pptx-angular-viewer` (et `/internals`, même fichier) et `pptx-vanilla-viewer` atteignent le package SDK du binding et `ooxml-ui/pptx/ai` par le même import dynamique que `/ai`. Aucun d'eux n'a le SDK `ai`, `@ai-sdk/*`, `zod` ou le catalogue d'automatisation dans ses chunks initiaux. Le handle ne change donc ni ce qui se charge, ni quand.

### Pourquoi le handle est livré maintenant {#why-the-handle-ships-now}

- Vous pouvez passer à `createAiAssistant` dans une version mineure, de sorte que la prochaine version majeure ne nécessite aucun changement de code de votre côté.
- Le handle garde l'assistant derrière un seul import auquel vous souscrivez. Dans React, Vue, Svelte et vanilla, le chargeur de panneau transporté par un handle vient de l'entrée `/ai`, et non de l'entrée principale. Dans Angular, le service de chat de la racine du package charge le runtime, de sorte que `/ai` ne transmet jamais de code à la racine.

### Ce que la prochaine version majeure change {#what-the-next-major-release-changes}

Les entrées principales deviennent exemptes du runtime de l'assistant lorsque toutes ces conditions sont réunies, dans chaque binding :

- Les graphes de modules source et construit des entrées principales (`pptx-react-viewer` et son `/no-i18next`, `pptx-vue-viewer` et `pptx-svelte-viewer` avec leurs entrées `/viewer`, `pptx-angular-viewer` et son `/internals`, `pptx-vanilla-viewer`) n'atteignent aucun module `ai`, `@ai-sdk/*`, `zod`, `ooxml-ui/pptx/ai` ou `ooxml-core/pptx/automation`, ni statiquement ni par un import dynamique.
- Leurs fichiers `.d.ts` n'atteignent aucun type de `ai` ou `@ai-sdk/*`.
- Le contrôle des packages de binding construits (`check-binding-packages.mjs`) attend que seule `./ai`, parmi les entrées de chaque binding, atteigne le package SDK du binding (`@ai-sdk/react`, `@ai-sdk/vue`, `@ai-sdk/svelte`) et `ooxml-ui/pptx/ai`.
- Une configuration simple passée à `ai` d'une entrée principale est retirée de son type et, à l'exécution, lève une erreur qui renvoie vers `createAiAssistant`.

Seule cette version change le comportement : `ai` n'accepte qu'un handle, les types de configuration `PptxAi*` (qui nomment des types du SDK `ai`) ne sont exportés que depuis les entrées `/ai`, et les entrées principales abandonnent leur propre chargeur paresseux de panneau. Le code qui passe déjà un handle continue de fonctionner sans changement. Dans Angular, le chargeur d'exécution se trouve aujourd'hui dans la racine du package : cette version doit donc aussi le déplacer dans `/ai`, sans ajouter d'export qui prend un chargeur.
