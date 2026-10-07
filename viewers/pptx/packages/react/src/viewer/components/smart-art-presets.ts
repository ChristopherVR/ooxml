/**
 * Thin re-export shim → `pptx-viewer-shared` (`render/smart-art-presets`).
 *
 * The SmartArt insert-gallery catalogue (presets + categories) now lives in
 * shared, consumed by every binding. This file preserves the historical React
 * import surface so `InsertSmartArtDialog.tsx` and the colocated test are
 * unchanged.
 */

export type { SmartArtCategory, SmartArtPreset } from 'ooxml-ui/pptx';
export { PRESETS, CATEGORIES } from 'ooxml-ui/pptx';
