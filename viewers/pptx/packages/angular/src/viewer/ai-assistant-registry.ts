/**
 * The registry of this package's `createAiAssistant` handles.
 *
 * The registry itself is private to `pptx-angular-viewer`: no entry exports it. ng-packagr
 * builds the `/ai` secondary entry point on its own and lets it import only package entries,
 * so `/ai` reaches this module through {@link createAiAssistantHandle}, exported from the
 * package root (and `/internals`, the same bundle). That function takes only a config, and a
 * handle carries nothing else: `AiChatService` loads the chat runtime itself, for a handle and
 * a plain config alike. `PowerPointViewerComponent` resolves its `ai` input here.
 */
import type { PptxAiConfig } from 'ooxml-ui/pptx/ai';
import { createAiAssistantRegistry } from 'ooxml-ui/pptx/ai/host';
import type { PptxAiAssistant } from 'ooxml-ui/pptx/ai/host';

export const aiAssistants = createAiAssistantRegistry<PptxAiConfig, undefined>({
	owner: 'pptx-angular-viewer',
});

/**
 * Backs `createAiAssistant` from `pptx-angular-viewer/ai`; import that instead. Not covered
 * by semver.
 *
 * @internal
 */
export function createAiAssistantHandle(config: PptxAiConfig): PptxAiAssistant {
	return aiAssistants.register(config, undefined);
}
