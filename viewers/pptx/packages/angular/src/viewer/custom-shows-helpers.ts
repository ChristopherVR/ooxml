/**
 * Thin re-export shim → vendored `pptx-viewer-shared` (`render/custom-shows`).
 *
 * The immutable custom-show list types/helpers were extracted to shared and are
 * consumed by every binding. This shim preserves the historical Angular import
 * surface.
 */

export type { CustomShow } from 'ooxml-ui/pptx';
export { generateCustomShowId, createCustomShow } from 'ooxml-ui/pptx';
