export { isCornerHandle, lockResizeAspect } from 'ooxml-core/geometry';

import { NUDGE_LARGE, NUDGE_SMALL, editorNudgeDelta } from './editor-keymap';

/**
 * PowerPoint compatibility surface: aspect-ratio geometry lives in core;
 * product-specific keyboard steps remain in the editor keymap.
 */

/**
 * Arrow-key nudge step in element px (Shift multiplies to the large step).
 * Aliases of the keymap's `NUDGE_SMALL`/`NUDGE_LARGE`: the step is defined once
 * so the inspector's position boxes and the keyboard cannot disagree.
 */
export const NUDGE_STEP = NUDGE_SMALL;
export const NUDGE_STEP_LARGE = NUDGE_LARGE;

/**
 * Map an arrow key to a nudge delta in element px, or `null` for other keys.
 * `large` (Shift held) uses the 10px step. Same function as the keymap's
 * `editorNudgeDelta`, kept under this name for the Svelte/Vanilla editors.
 */
export const nudgeDelta: typeof editorNudgeDelta = editorNudgeDelta;
