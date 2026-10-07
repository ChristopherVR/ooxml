/** Compatibility seam for the shared, framework-agnostic AI message helpers. */
export { extractReadyToolCalls, toRenderableParts } from 'ooxml-ui/pptx/ai';
export type {
	AiUiMessage,
	ReadyToolCall,
	RenderablePart,
	RenderableTextPart,
	RenderableToolPart,
} from 'ooxml-ui/pptx/ai';
