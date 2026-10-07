/**
 * Thin re-export shim -> vendored `pptx-viewer-shared`.
 *
 * The hyperlink-dialog patch-builders were extracted to `pptx-viewer-shared`
 * (`render/hyperlink-dialog.ts`) and are consumed by every binding. This shim
 * preserves the historical Angular import surface so `HyperlinkDialogComponent`,
 * the viewer barrel, and the colocated tests are unchanged.
 */
export type { HyperlinkDraft } from 'ooxml-ui/pptx';
export {
	hasExistingLink,
	seedHyperlinkDraft,
	buildHyperlinkPatch,
	buildClearHyperlinkPatch,
} from 'ooxml-ui/pptx';
