// Excel-for-the-web sheet tab bar: scroll arrows and "+" on the left, then the tab strip.
export const TABS_CSS = `
.xst {
	display: flex;
	align-items: stretch;
	height: 32px;
	box-sizing: border-box;
	border-top: 1px solid var(--xve-border, #e1dfdd);
	background: var(--xve-secondary, #f3f2f1);
	color: var(--xve-foreground, #242424);
	font: 13px/1 'Segoe UI', system-ui, -apple-system, sans-serif;
	user-select: none;
	overflow: hidden;
}
.xst-nav {
	display: flex;
	align-items: center;
	gap: 2px;
	padding: 0 4px;
	flex: none;
}
.xst-btn {
	display: inline-flex;
	align-items: center;
	justify-content: center;
	width: 28px;
	height: 26px;
	padding: 0;
	border: 0;
	border-radius: 4px;
	background: transparent;
	color: inherit;
	font: inherit;
	font-size: 14px;
	cursor: pointer;
}
.xst-btn:hover:not(:disabled) { background: var(--xve-accent, #e1dfdd); }
.xst-btn:disabled { color: var(--xve-muted-foreground, #a19f9d); cursor: default; }
.xst-btn:focus-visible, .xst-tab:focus-visible {
	outline: 2px solid var(--xve-ring, #616161);
	outline-offset: -2px;
}
.xst-btn[hidden] { display: none; }
.xst-strip {
	position: relative;
	display: flex;
	align-items: stretch;
	flex: 1;
	min-width: 0;
	overflow: hidden;
	scroll-behavior: smooth;
}
.xst-tab {
	position: relative;
	display: inline-flex;
	align-items: center;
	flex: none;
	max-width: 240px;
	padding: 0 14px;
	border: 0;
	border-right: 1px solid var(--xve-border, #e1dfdd);
	background: transparent;
	color: inherit;
	font: inherit;
	white-space: nowrap;
	cursor: pointer;
}
.xst-tab:hover:not(.xst-active) { background: var(--xve-accent, #e9e8e7); }
.xst-tab .xst-label { overflow: hidden; text-overflow: ellipsis; }
.xst-tab::after {
	content: '';
	position: absolute;
	left: 8px;
	right: 8px;
	bottom: 2px;
	height: 3px;
	border-radius: 2px;
	background: var(--xst-color, transparent);
}
.xst-tab.xst-colored:not(.xst-active) {
	background: color-mix(in srgb, var(--xst-color) 22%, transparent);
}
.xst-tab.xst-active {
	background: var(--xve-sheet-bg, #ffffff);
	color: var(--xve-selection-border, #217346);
	font-weight: 600;
}
.xst-tab.xst-active::after {
	background: var(--xve-selection-border, #217346);
}
.xst-tab.xst-active.xst-colored::after { background: var(--xst-color); }
.xst-tab.xst-dragging { opacity: 0.6; }
.xst-rename {
	width: 120px;
	height: 22px;
	box-sizing: border-box;
	padding: 0 4px;
	border: 1px solid var(--xve-selection-border, #217346);
	border-radius: 2px;
	font: inherit;
	color: var(--xve-foreground, #242424);
	background: var(--xve-sheet-bg, #ffffff);
}
.xst-marker {
	position: absolute;
	top: 4px;
	bottom: 4px;
	width: 2px;
	margin-left: -1px;
	background: var(--xve-selection-border, #217346);
	pointer-events: none;
}
`;
