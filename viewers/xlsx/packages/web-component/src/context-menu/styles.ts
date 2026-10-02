// Office-for-the-web menu look: rounded card, soft shadow, 32px rows, right-aligned shortcuts.
export const MENU_CSS = `
.xcm-menu {
	position: fixed;
	z-index: 1000;
	min-width: 220px;
	max-height: calc(100vh - 8px);
	overflow-y: auto;
	box-sizing: border-box;
	padding: 4px 0;
	border: 1px solid var(--xve-border, #e1dfdd);
	border-radius: 6px;
	background: var(--xve-popover, #ffffff);
	color: var(--xve-popover-foreground, #242424);
	box-shadow: 0 8px 16px rgba(0, 0, 0, 0.14), 0 0 2px rgba(0, 0, 0, 0.12);
	font: 13px/1.2 'Segoe UI', system-ui, -apple-system, sans-serif;
	user-select: none;
}
.xcm-item {
	display: flex;
	align-items: center;
	gap: 8px;
	width: 100%;
	height: 32px;
	box-sizing: border-box;
	padding: 0 16px 0 8px;
	border: 0;
	background: transparent;
	color: inherit;
	font: inherit;
	text-align: left;
	cursor: default;
	outline: none;
}
.xcm-item:focus,
.xcm-item:hover:not([aria-disabled]) {
	background: var(--xve-accent, #f3f2f1);
}
.xcm-item:focus-visible {
	box-shadow: inset 0 0 0 1px var(--xve-ring, #616161);
}
.xcm-item[aria-disabled] {
	color: var(--xve-muted-foreground, #a19f9d);
}
.xcm-check {
	width: 16px;
	flex: none;
	text-align: center;
}
.xcm-label {
	flex: 1;
	white-space: nowrap;
}
.xcm-shortcut {
	margin-left: 24px;
	color: var(--xve-muted-foreground, #616161);
	white-space: nowrap;
}
.xcm-separator {
	height: 1px;
	margin: 4px 0;
	background: var(--xve-border, #e1dfdd);
}
`;
