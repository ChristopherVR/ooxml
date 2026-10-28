/**
 * `pptx-angular-viewer/ai`: the AI assistant entry (an ng-packagr secondary entry point,
 * built to its own FESM file).
 *
 * Importing this entry is the opt-in to the chat runtime and its optional `ai` peer. The
 * runtime itself still loads lazily, the first time the user opens the assistant.
 *
 * ```ts
 * import { createAiAssistant } from 'pptx-angular-viewer/ai';
 *
 * readonly assistant = createAiAssistant({ connection: { kind: 'endpoint', api: '/api/chat' } });
 * // <pptx-viewer [content]="bytes" [ai]="assistant" />
 * ```
 */
import type { PptxAiConfig } from 'ooxml-ui/pptx/ai';
import type { PptxAiAssistant } from 'ooxml-ui/pptx/ai/host';
// ng-packagr builds this entry point on its own, so it reaches the viewer's handle registry
// through the package root, the module the viewer itself resolves `ai` with.
import { createAiAssistantHandle } from 'pptx-angular-viewer';

/**
 * Create an AI assistant for the viewer's `ai` input from a `PptxAiConfig`.
 *
 * The handle is opaque and loads nothing. It keeps `config` by reference; create a new
 * handle to change it. The viewer's chat panel loads the chat runtime, for a handle and a
 * plain config alike.
 */
export function createAiAssistant(config: PptxAiConfig): PptxAiAssistant {
	return createAiAssistantHandle(config);
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
