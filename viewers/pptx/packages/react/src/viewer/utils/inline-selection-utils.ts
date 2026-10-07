/**
 * Thin re-export shim. The framework-agnostic inline-selection helpers now live
 * in `pptx-viewer-shared`.
 */
export {
	setPendingSelectionRestore,
	getPendingSelectionRestore,
	getInlineEditorSelection,
	applyStyleToSelectedSegments,
	restoreSegmentSelection,
} from 'ooxml-ui/pptx';
export type { InlineTextSelection } from 'ooxml-ui/pptx';
