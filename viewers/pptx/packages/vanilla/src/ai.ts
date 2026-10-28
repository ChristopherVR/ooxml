/**
 * `pptx-vanilla-viewer/ai`: the AI assistant entry.
 *
 * Importing this entry is the opt-in to the chat panel and its optional `ai` peer. The
 * panel itself still loads lazily, the first time the user opens it.
 *
 * ```ts
 * import { createPptxViewer } from 'pptx-vanilla-viewer';
 * import { createAiAssistant } from 'pptx-vanilla-viewer/ai';
 *
 * const ai = createAiAssistant({ connection: { kind: 'endpoint', api: '/api/chat' } });
 * const viewer = createPptxViewer(container, { ai });
 * ```
 */
import type { PptxAiConfig } from 'ooxml-ui/pptx/ai';
import type { PptxAiAssistant } from 'ooxml-ui/pptx/ai/host';

import { aiAssistants } from './viewer/ai-assistant-registry';
import type { AiPanelLoader } from './viewer/ai/ai-panel-loader';

const loadPanel: AiPanelLoader = () => import('./viewer/ai/ai-panel');

/**
 * Create an AI assistant for the viewer's `ai` option from a `PptxAiConfig`.
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
