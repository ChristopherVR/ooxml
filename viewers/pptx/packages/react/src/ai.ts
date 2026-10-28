/**
 * `pptx-react-viewer/ai`: the AI assistant entry.
 *
 * Importing this entry is the opt-in to the chat panel and its optional `ai` and
 * `@ai-sdk/react` peers. The panel itself still loads lazily, the first time the user
 * opens it.
 *
 * ```tsx
 * import { createAiAssistant } from 'pptx-react-viewer/ai';
 *
 * const assistant = createAiAssistant({ connection: { kind: 'endpoint', api: '/api/chat' } });
 * <PowerPointViewer content={bytes} ai={assistant} />;
 * ```
 */
import type { PptxAiConfig } from 'ooxml-ui/pptx/ai';
import type { PptxAiAssistant } from 'ooxml-ui/pptx/ai/host';

import { aiAssistants } from './viewer/ai-assistant-registry';
import { AiChatPanelLazy } from './viewer/components/ai/AiChatPanelLazy';

/**
 * Create an AI assistant for the viewer's `ai` prop from a `PptxAiConfig`.
 *
 * The handle is opaque and loads nothing. It keeps `config` by reference; create a new
 * handle to change it. Every handle renders the same chat panel component, so a handle
 * created on every render (rather than in `useMemo` or at module scope) never remounts the
 * chat.
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
