/**
 * `pptx-vue-viewer/ai`: the AI assistant entry.
 *
 * Importing this entry is the opt-in to the chat panel and its optional `ai` and
 * `@ai-sdk/vue` peers. The panel itself still loads lazily (an async component), the first
 * time the user opens it.
 *
 * ```vue
 * <script setup lang="ts">
 * import { PowerPointViewer } from 'pptx-vue-viewer';
 * import { createAiAssistant } from 'pptx-vue-viewer/ai';
 *
 * const assistant = createAiAssistant({ connection: { kind: 'endpoint', api: '/api/chat' } });
 * </script>
 *
 * <template><PowerPointViewer :content="bytes" :ai="assistant" /></template>
 * ```
 */
import type { PptxAiConfig } from 'ooxml-ui/pptx/ai';
import type { PptxAiAssistant } from 'ooxml-ui/pptx/ai/host';

import { aiAssistants } from './viewer/ai-assistant-registry';
import { AiChatPanelLazy } from './viewer/components/ai';

/**
 * Create an AI assistant for the viewer's `ai` prop from a `PptxAiConfig`.
 *
 * The handle is opaque and loads nothing. It keeps `config` by reference; create a new
 * handle to change it. Every handle carries the same async chat panel component, so
 * replacing the handle never remounts the chat.
 */
export function createAiAssistant(config: PptxAiConfig): PptxAiAssistant {
	return aiAssistants.register(config, AiChatPanelLazy);
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
