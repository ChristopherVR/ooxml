import { FOCUS_RING, tok } from './tokens.js';

/**
 * `office-ui-select` styles, moved from pptx-viewer (`select-styles.ts`, `select-ribbon-styles.ts`)
 * and expressed in tokens. `variant="ribbon-font"` and `variant="ribbon-icon"` are the compact
 * ribbon combo and icon picker; `data-font-picker="family"` widens and lengthens the popup.
 */
export const SELECT_CSS = `
:host { display: inline-block; box-sizing: border-box; min-width: 0; max-width: 100%;
	color: ${tok('--office-foreground')}; font: inherit; font-size: ${tok('--office-font-size-sm')}; vertical-align: middle; }
:host([disabled]) { opacity: .5; cursor: not-allowed; }
button { box-sizing: border-box; width: 100%; min-height: ${tok('--office-field-height')}; display: flex; align-items: center;
	justify-content: space-between; gap: ${tok('--office-space-2-5')}; padding: ${tok('--office-space-1')} ${tok('--office-space-1-5')};
	border: ${tok('--office-border-width')} solid ${tok('--office-field-border')}; border-radius: ${tok('--office-field-radius')};
	background: ${tok('--office-select-background')}; color: inherit; font: inherit; font-size: inherit; text-align: start; cursor: pointer; }
button:focus-visible { ${FOCUS_RING} }
button:disabled { cursor: not-allowed; }
.value { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.chevron svg { display: none; }
.chevron { flex: none; width: ${tok('--office-chevron-size')}; height: ${tok('--office-chevron-size')};
	margin: 0 ${tok('--office-space-0')} ${tok('--office-space-0')} 0;
	border-right: calc(${tok('--office-border-width')} * 1.5) solid currentColor;
	border-bottom: calc(${tok('--office-border-width')} * 1.5) solid currentColor; transform: rotate(45deg); }
:host([open]) .chevron { margin: ${tok('--office-space-0')} ${tok('--office-space-0')} 0 0; transform: rotate(225deg); }
.menu { box-sizing: border-box; position: fixed; margin: 0; padding: ${tok('--office-space-1')};
	min-width: ${tok('--office-select-menu-min-width')}; max-height: min(${tok('--office-select-max-height')}, 50vh); overflow: auto;
	border: ${tok('--office-border-width')} solid ${tok('--office-field-border')}; border-radius: ${tok('--office-field-radius')};
	background: ${tok('--office-popover')}; color: ${tok('--office-popover-foreground')}; box-shadow: ${tok('--office-shadow')};
	font: inherit; font-size: inherit; }
.menu:not(:popover-open):not([data-fallback-open]) { display: none; }
.option { display: block; box-sizing: border-box; width: 100%; padding: ${tok('--office-space-1-5')} ${tok('--office-space-2')};
	border-radius: ${tok('--office-radius-sm')}; cursor: pointer; white-space: nowrap; }
.group { padding: ${tok('--office-space-1-5')} ${tok('--office-space-2')} ${tok('--office-space-0')};
	color: ${tok('--office-muted-foreground')}; font-size: .9em; font-weight: ${tok('--office-font-weight-bold')}; }
.option:not([aria-disabled="true"]):hover, .option[data-active]:not([aria-selected="true"]) {
	background: ${tok('--office-select-option-active')}; }
.option[aria-selected="true"] { background: ${tok('--office-accent')}; color: ${tok('--office-accent-foreground')}; }
.option[aria-selected="true"]:hover { filter: brightness(1.08); }
.option[aria-disabled="true"] { opacity: .45; cursor: not-allowed; }
.description { flex: none; color: ${tok('--office-muted-foreground')}; font-size: ${tok('--office-font-size-2xs')}; }

:host([variant^="ribbon-"][disabled]) { opacity: .4; }
:host([variant^="ribbon-"]) button { height: ${tok('--office-ribbon-combo-height')}; min-height: ${tok('--office-ribbon-combo-height')};
	padding: ${tok('--office-space-1')} ${tok('--office-space-2')}; gap: ${tok('--office-space-1')};
	border: ${tok('--office-border-width')} solid color-mix(in oklab, ${tok('--office-border')} 60%, transparent);
	border-radius: ${tok('--office-radius-sm')}; font-size: ${tok('--office-font-size-sm')};
	background: color-mix(in oklab, ${tok('--office-background')} 60%, transparent); color: ${tok('--office-foreground')}; }
:host([variant^="ribbon-"]) .value { flex: 1; min-width: 0; }
:host([variant^="ribbon-"]) .chevron { width: ${tok('--office-icon-size-md')}; height: ${tok('--office-icon-size-md')}; margin: 0; border: 0; transform: none; }
:host([variant^="ribbon-"]) .chevron svg { display: block; width: ${tok('--office-icon-size-md')}; height: ${tok('--office-icon-size-md')}; }
:host([variant="ribbon-font"]) .chevron { color: ${tok('--office-muted-foreground')}; }
:host([variant="ribbon-font"][data-font-picker="size"]) button { padding-inline: ${tok('--office-space-1-5')}; }
:host([variant="ribbon-icon"]) button { width: ${tok('--office-ribbon-icon-select-width')}; gap: ${tok('--office-space-1-5')};
	padding: ${tok('--office-space-1-5')} ${tok('--office-space-2-5')}; border: 0; background: ${tok('--office-surface')}; justify-content: center; }
:host([variant="ribbon-icon"]) :is(.value, .chevron) { display: none; }
::slotted(svg[slot="icon"]) { display: block; width: ${tok('--office-icon-size-md')}; height: ${tok('--office-icon-size-md')}; flex: none; }
:host([variant^="ribbon-"]) button:hover:not(:disabled) { background: ${tok('--office-selected')}; }
:host([variant^="ribbon-"]) .menu { width: ${tok('--office-select-ribbon-menu-width')}; padding: ${tok('--office-space-1')} 0;
	border-radius: ${tok('--office-radius-lg')}; font-size: ${tok('--office-font-size-sm')}; }
:host([data-font-picker="family"]) .menu { width: ${tok('--office-font-picker-width')}; }
:host([variant^="ribbon-"]) .option { display: flex; justify-content: space-between; gap: ${tok('--office-space-3')};
	padding: ${tok('--office-space-1-5')} ${tok('--office-space-3')}; border-radius: 0; }
:host([variant^="ribbon-"]) .option[aria-selected="true"] { background: transparent; color: ${tok('--office-foreground')}; }
:host([variant^="ribbon-"]) .option:hover, :host([variant^="ribbon-"]) .option[data-active] { background: ${tok('--office-surface')}; filter: none; }
:host([variant^="ribbon-"]) .group { padding: ${tok('--office-space-2')} ${tok('--office-space-3')} ${tok('--office-space-1')};
	text-transform: uppercase; letter-spacing: .02em; font-size: ${tok('--office-font-size-2xs')}; }
@media (pointer: coarse), (max-width: 767px) {
	button, :host([variant^="ribbon-"]) button { min-height: ${tok('--office-target-size-touch')}; }
	:host([variant^="ribbon-"]) button { height: ${tok('--office-target-size-touch')}; }
}
@media (forced-colors: active) {
	button, .menu { border-color: CanvasText; background: Canvas; color: CanvasText; }
	.option[aria-selected="true"], .option[data-active], .option:not([aria-disabled="true"]):hover {
		background: Highlight; color: HighlightText; }
}
`;
