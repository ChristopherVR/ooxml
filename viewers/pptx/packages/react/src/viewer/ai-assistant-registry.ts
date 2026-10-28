/**
 * The registry of this package's `createAiAssistant` handles.
 *
 * Private to `pptx-react-viewer`: no package entry exports it. The `/ai` entry's
 * `createAiAssistant` registers handles here and `ViewerSidePanels` resolves the viewer's `ai`
 * prop here, so the bundle shares this module between the main entries and `/ai` and both
 * see one registry. A handle from anywhere else is read as a plain config.
 */
import type { PptxAiConfig } from 'ooxml-ui/pptx/ai';
import { createAiAssistantRegistry } from 'ooxml-ui/pptx/ai/host';

import type { AiChatPanelLazy } from './components/ai/AiChatPanelLazy';

export const aiAssistants = createAiAssistantRegistry<PptxAiConfig, typeof AiChatPanelLazy>({
	owner: 'pptx-react-viewer',
});
