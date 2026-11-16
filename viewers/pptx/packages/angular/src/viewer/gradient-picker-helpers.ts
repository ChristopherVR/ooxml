/**
 * Thin re-export shim -> vendored `pptx-viewer-shared`.
 *
 * The pure gradient-picker readers + patch-builders were extracted to
 * `pptx-viewer-shared` (`render/gradient-picker.ts`). This shim preserves the
 * historical Angular import surface (`internals.ts`); the inspector's own
 * gradient stop list now runs on the shared fill actions (`shapeFillPatch`).
 */
export type { GradientStop, GradientState } from 'ooxml-ui/pptx';
export {
	gradientStateOf,
	gradientStateFromStyle,
	hasGradientFill,
	gradientStatePatch,
	addGradientStopPatch,
	removeGradientStopPatch,
	updateGradientStopPatch,
	gradientStopColorCommitPatch,
} from 'ooxml-ui/pptx';
