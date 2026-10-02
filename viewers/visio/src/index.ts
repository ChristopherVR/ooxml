export { ViewerController, type ViewerState } from './controller.js';
export { mountViewer, type MountedViewer } from './binding.js';
export { VisioViewerElement, registerVisioViewer } from './viewer-element.js';
export { renderPage, type RenderResult } from './render-svg.js';
export {
	eventKeys,
	propertyKeys,
	type ViewerProperties,
	type ViewerEvents,
	type ViewerCallbacks,
	type ViewerOptions,
	type VsdxSource,
} from './contract.js';
export type { VisioDocument, VisioPage, VisioShape, VisioDiagnostic } from 'ooxml-core/visio';

export { createWorkerParser, type CancellableParser } from './worker-parser.js';

export { compatibilityNotes, compatibilityText, type CompatibilityNote } from './diagnostics.js';
