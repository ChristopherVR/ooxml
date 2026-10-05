export { ViewerController, type ViewerState } from './controller.js';
export { mountViewer, type MountedViewer } from './binding.js';
export { VisioViewerElement, registerVisioViewer } from './viewer-element.js';
export { renderPage, type RenderResult, type RenderOptions } from './render-svg.js';
export {
	exportPageSvg,
	MAX_SVG_EXPORT_BYTES,
	type SvgExportOptions,
	type SvgExportResult,
} from './export-svg.js';
export {
	createPrintSnapshot,
	PRINT_SNAPSHOT_LIMITS,
	type PrintSnapshot,
	type PrintSnapshotPage,
	type PrintSnapshotOptions,
	type PrintSnapshotLimits,
	type PrintSnapshotUsage,
	type CurrentPagePrintSnapshotOptions,
} from './print-snapshot.js';
export {
	eventKeys,
	propertyKeys,
	type ViewerProperties,
	type ViewerEvents,
	type ViewerCallbacks,
	type ViewerOptions,
	type VsdxSource,
} from 'ooxml-core/visio/ui';
export type { VisioDocument, VisioPage, VisioShape, VisioDiagnostic } from 'ooxml-core/visio';

export { createWorkerParser, type CancellableParser } from './worker-parser.js';

export { compatibilityNotes, compatibilityText, type CompatibilityNote } from 'ooxml-core/visio/ui';

export {
	TEXT_SEARCH_LIMITS,
	type TextSearchResult,
	type TextSearchState,
} from 'ooxml-core/visio/ui';

export { VIEWER_LAYER_LIMITS, type LayerVisibilityOverride } from './viewer-layers.js';

export type { ViewerEditState, VsdxExportResult } from 'ooxml-core/visio/ui';

export type { VisioEdit, VisioGeometryEdit } from 'ooxml-core/visio';

/** The built-in demo document, for the demo and tests. */
export { demoDocument } from 'ooxml-core/visio/ui';
