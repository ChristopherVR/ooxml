/**
 * `pptx-svelte-viewer/ai`: the AI assistant entry.
 *
 * Importing this entry is the opt-in to the chat panel and its optional `ai` and
 * `@ai-sdk/svelte` peers. The panel itself still loads lazily, the first time the user
 * opens it.
 *
 * ```svelte
 * <script lang="ts">
 * 	import { PowerPointViewer } from 'pptx-svelte-viewer';
 * 	import { createAiAssistant } from 'pptx-svelte-viewer/ai';
 *
 * 	const assistant = createAiAssistant({ connection: { kind: 'endpoint', api: '/api/chat' } });
 * </script>
 *
 * <PowerPointViewer content={bytes} ai={assistant} />
 * ```
 */
import type { PptxAiConfig } from 'ooxml-ui/pptx/ai';
import type { PptxAiAssistant } from 'ooxml-ui/pptx/ai/host';

import { aiAssistants } from './viewer/ai-assistant-registry';
import type { AiChatPanelLoader } from './viewer/components/ai/ai-chat-panel-loader';

const loadPanel: AiChatPanelLoader = () => import('./viewer/components/ai/AiChatPanel.svelte');

/**
 * Create an AI assistant for the viewer's `ai` prop from a `PptxAiConfig`.
 *
 * The handle is opaque and loads nothing. It keeps `config` by reference; create a new
 * handle to change it. Every handle carries the same panel loader.
 */
export function createAiAssistant(config: PptxAiConfig): PptxAiAssistant {
	return aiAssistants.register(config, loadPanel);
}

export type { PptxAiAssistant } from 'ooxml-ui/pptx/ai/host';
export type {
	PptxAiBridge,
	PptxAiConfig,
	PptxAiConnection,
	PptxAiContextStrategy,
	PptxAiToolName,
	PptxAiWritePolicy,
} from 'ooxml-ui/pptx/ai';
