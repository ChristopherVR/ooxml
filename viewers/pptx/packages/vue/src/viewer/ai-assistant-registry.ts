/**
 * The registry of this package's `createAiAssistant` handles.
 *
 * Private to `pptx-vue-viewer`: no package entry exports it. The `/ai` entry's
 * `createAiAssistant` registers handles here and `useViewerCustomization` resolves the
 * viewer's `ai` prop here, so the build shares this module between the main entries and `/ai`
 * and both see one registry. A handle from anywhere else is read as a plain config.
 */
import type { PptxAiConfig } from 'ooxml-ui/pptx/ai';
import { createAiAssistantRegistry } from 'ooxml-ui/pptx/ai/host';
import type { Component } from 'vue';

export const aiAssistants = createAiAssistantRegistry<PptxAiConfig, Component>({
	owner: 'pptx-vue-viewer',
});
