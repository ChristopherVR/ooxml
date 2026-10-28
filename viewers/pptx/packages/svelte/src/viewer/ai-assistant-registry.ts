/**
 * The registry of this package's `createAiAssistant` handles.
 *
 * Private to `pptx-svelte-viewer`: no package entry exports it. The `/ai` entry's
 * `createAiAssistant` registers handles here and `PowerPointViewer.svelte` resolves the
 * viewer's `ai` prop here, so the build shares this module between the main entries and `/ai`
 * and both see one registry. A handle from anywhere else is read as a plain config.
 */
import { createAiAssistantRegistry } from 'ooxml-ui/pptx/ai/host';

import type { AiChatPanelLoader } from './components/ai/ai-chat-panel-loader';
import type { PptxAiConfig } from './types';

export const aiAssistants = createAiAssistantRegistry<PptxAiConfig, AiChatPanelLoader>({
	owner: 'pptx-svelte-viewer',
});
