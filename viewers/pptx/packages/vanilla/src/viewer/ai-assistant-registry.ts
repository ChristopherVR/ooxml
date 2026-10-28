/**
 * The registry of this package's `createAiAssistant` handles.
 *
 * Private to `pptx-vanilla-viewer`: no package entry exports it. The `/ai` entry's
 * `createAiAssistant` registers handles here and `PptxViewer` resolves the viewer's `ai`
 * option here, so the build shares this module between the main entries and `/ai` and both
 * see one registry. A handle from anywhere else is read as a plain config.
 */
import type { PptxAiConfig } from 'ooxml-ui/pptx/ai';
import { createAiAssistantRegistry } from 'ooxml-ui/pptx/ai/host';

import type { AiPanelLoader } from './ai/ai-panel-loader';

export const aiAssistants = createAiAssistantRegistry<PptxAiConfig, AiPanelLoader>({
	owner: 'pptx-vanilla-viewer',
});
