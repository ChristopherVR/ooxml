/**
 * Thin re-export shim -> `pptx-viewer-shared` (via the vendored copy).
 *
 * The EyeDropper colour sampler (native API + DOM-sampling fallback) now lives
 * in `pptx-viewer-shared` (`render/eyedropper`). This shim preserves the
 * historical import surface for Angular's format-painter service.
 */
export type { EyedropperResult } from 'ooxml-ui/pptx';
export {
	eyedropperAvailable,
	openNativeEyeDropper,
	pickColorByClickFallback,
	sampleColorFromSlide,
} from 'ooxml-ui/pptx';
