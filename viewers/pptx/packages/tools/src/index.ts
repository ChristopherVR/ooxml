export type { ToolContext, ToolResult } from './types';
export type {
	CollaborationProvider,
	FileSystemProvider,
	ViewerProvider,
	ExecutionContext,
} from './types';
export { loadPresentation, savePresentation, executeToolWithContext } from './execution';
export * from './tools/index';

// Re-export the core engine so consumers can load/save PPTX files without a
// separate `ooxml-core` install. The engine ships as a dependency, so a
// single `npm install pptx-viewer-mcp` is enough to use the tools end to end.
export { PptxHandler } from 'ooxml-core/pptx';
export type { PptxData, PptxSlide, PptxElement } from 'ooxml-core/pptx';
