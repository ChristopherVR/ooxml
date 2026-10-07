/**
 * Thin re-export shim → `pptx-viewer-shared`.
 *
 * Snap-to-grid geometry (View ▸ Snap to Grid) was consolidated into
 * `pptx-viewer-shared` (`render/snap-guides`). This shim preserves the
 * historical Vue import surface (`snapValue`, `snapBox`, `SnapBox`) so
 * `PowerPointViewer.vue` and the colocated test are unchanged.
 */

export type { SnapBox } from 'ooxml-ui/pptx';

export { snapValue, snapBox } from 'ooxml-ui/pptx';
