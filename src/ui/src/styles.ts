import { TOUCH, tok } from './tokens.js';

/**
 * Rules every control repeats: `hidden`, the coarse-pointer touch target and the forced-colors
 * focus ring. `controlStyles` (base.ts) puts it ahead of each element's own `.css` file.
 */
export const COMMON_CONTROL_CSS = `
:host([hidden]) { display: none !important; }
@media ${TOUCH} {
	:host { --office-target-size: ${tok('--office-target-size-touch')}; }
}
@media (forced-colors: active) {
	:host(:focus-visible), :focus-visible { outline: ${tok('--office-focus-width')} solid Highlight; outline-offset: ${tok('--office-focus-offset')}; }
}
`;
