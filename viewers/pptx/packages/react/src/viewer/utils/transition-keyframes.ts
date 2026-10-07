/**
 * Thin re-export shim: the slide-transition `@keyframes` and result type now
 * live in `pptx-viewer-shared` (`SLIDE_TRANSITION_KEYFRAMES`). Kept so existing
 * importers (`slide-transitions` barrel, overlay, tests) resolve unchanged.
 */
export type { SlideTransitionAnimations } from 'ooxml-ui/pptx';
export { SLIDE_TRANSITION_KEYFRAMES } from 'ooxml-ui/pptx';
