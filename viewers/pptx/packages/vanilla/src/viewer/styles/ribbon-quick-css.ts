/**
 * Quick-access chrome stylesheet fragment: the primary row's React-aligned
 * right cluster (Present split button, popover menus, "+ Show") and the tab
 * row's right-side actions (Record, Share). Split from `ribbon-css.ts` to
 * keep both files inside the repo's per-file size budget; concatenated right
 * after it by {@link buildViewerCss}.
 */
export const RIBBON_QUICK_CSS = `
/* ── Primary-row quick actions ──────────────────────────────────────────── */
.pptxv-primary-sep { width: 1px; align-self: stretch; margin: 0 4px; background: color-mix(in srgb, var(--pptx-border) 40%, transparent); }
/* React-style "Present" split button + its options dropdown. */
.pptxv-present-split { position: relative; display: inline-flex; align-items: stretch; border-left: 1px solid var(--pptx-border); }
.pptxv-present-main {
	padding: 2px 8px;
	border: none;
	background: transparent;
	color: var(--pptx-foreground);
	font: inherit;
	font-size: 12px;
	cursor: pointer;
}
.pptxv-present-main:hover { background: var(--pptx-accent); }
.pptxv-present-caret {
	display: inline-flex;
	align-items: center;
	padding: 0 4px;
	border: none;
	border-left: 1px solid var(--pptx-border);
	background: transparent;
	color: var(--pptx-foreground);
	cursor: pointer;
}
.pptxv-present-caret:hover, .pptxv-present-caret.is-active { background: var(--pptx-accent); }
.pptxv-present-caret svg { width: 12px; height: 12px; }
/* Right-aligned popover menus (Present options + "..." overflow). */
.pptxv-primary-menu {
	position: absolute;
	top: calc(100% + 4px);
	right: 0;
	z-index: 40;
	min-width: 190px;
	max-height: 320px;
	overflow-y: auto;
	padding: 4px;
	border: 1px solid var(--pptx-border);
	border-radius: var(--pptx-radius);
	background: var(--pptx-card);
	color: var(--pptx-card-foreground);
	box-shadow: 0 6px 20px rgb(0 0 0 / 0.25);
}
.pptxv-primary-menu[hidden] { display: none; }
.pptxv-primary-menu-item {
	display: block;
	width: 100%;
	padding: 6px 8px;
	border: none;
	border-radius: 4px;
	background: transparent;
	color: inherit;
	font: inherit;
	font-size: 12px;
	text-align: left;
	white-space: nowrap;
	cursor: pointer;
}
.pptxv-primary-menu-item:hover { background: var(--pptx-accent); color: var(--pptx-accent-foreground); }
.pptxv-primary-menu-sep { margin: 4px 0; border-top: 1px solid color-mix(in srgb, var(--pptx-border) 60%, transparent); }
.pptxv-primary-menu-host { position: relative; display: inline-flex; }
/* Custom-shows quick action ("+ Show"). */
.pptxv-show-btn {
	padding: 2px 8px;
	border: none;
	border-radius: var(--pptx-radius);
	background: var(--pptx-muted);
	color: var(--pptx-foreground);
	font: inherit;
	font-size: 11px;
	white-space: nowrap;
	cursor: pointer;
}
.pptxv-show-btn:hover { background: var(--pptx-accent); }
.pptxv-read-only-badge {
	display: inline-flex;
	align-items: center;
	padding: 1px 8px;
	border-radius: 3px;
	background: color-mix(in srgb, #d97706 90%, transparent);
	color: #fffbeb;
	font-size: 10px;
	white-space: nowrap;
}
.pptxv-read-only-badge[hidden] { display: none; }
/* ── Tab-row right-side actions ── */
/* Right side of the tab row: Record + Share, matching React's TabRowActions. */
.pptxv-tabrow-spacer { flex: 1; }
.pptxv-tabrow-actions { display: inline-flex; align-items: center; gap: 4px; padding-right: 4px; }
.pptxv-tabrow-record {
	display: inline-flex;
	align-items: center;
	gap: 5px;
	padding: 3px 9px;
	border: none;
	border-radius: 3px;
	background: transparent;
	color: var(--pptx-muted-foreground);
	font: inherit;
	font-size: 11px;
	font-weight: 500;
	white-space: nowrap;
	cursor: pointer;
}
.pptxv-tabrow-record:hover { background: var(--pptx-accent); color: var(--pptx-foreground); }
.pptxv-tabrow-record-dot { width: 8px; height: 8px; border-radius: 50%; background: #ef4444; }
/* ── Ribbon Display Options (Always show Ribbon / Show tabs only / Full-screen mode) ── */
.pptxv-ribbon-display-btn {
	display: inline-flex;
	align-items: center;
	justify-content: center;
	width: 22px;
	height: 22px;
	padding: 0;
	border: none;
	border-radius: 3px;
	background: transparent;
	color: var(--pptx-muted-foreground);
	cursor: pointer;
}
.pptxv-ribbon-display-btn:hover, .pptxv-ribbon-display-btn.is-active { background: var(--pptx-accent); color: var(--pptx-foreground); }
/* Collapsed to the tab row, the chevron points the other way, as PowerPoint's pin does. */
.pptxv-ribbon[data-ribbon-display='tabs'] .pptxv-ribbon-display-btn svg { transform: rotate(180deg); }
/* The menu hangs under the right end of the tab row; the row itself scrolls and would clip it. */
.pptxv-ribbon-display-anchor { position: relative; height: 0; }
.pptxv-ribbon-display-menu { top: 2px; right: 6px; }
.pptxv-ribbon-display-menu .pptxv-primary-menu-item { padding-left: 26px; position: relative; }
.pptxv-ribbon-display-menu .pptxv-primary-menu-item[aria-checked='true']::before {
	content: '';
	position: absolute;
	left: 9px;
	top: 50%;
	width: 9px;
	height: 5px;
	border-left: 1.5px solid currentColor;
	border-bottom: 1.5px solid currentColor;
	transform: translateY(-70%) rotate(-45deg);
}
/* "Show tabs only": a chosen tab lays its commands over the slide instead of pushing it down. */
.pptxv-ribbon > [data-ribbon-overlay] {
	position: absolute;
	left: 0;
	right: 0;
	top: var(--pptxv-ribbon-overlay-top, 0);
	z-index: 35;
	border-bottom: 1px solid var(--pptx-border);
	/* Opaque: some themes make the card colour translucent, and this lies over the slide. */
	background: linear-gradient(var(--pptx-card), var(--pptx-card)), var(--pptx-background);
	box-shadow: 0 8px 20px rgb(0 0 0 / 0.2);
}
/* Full-screen mode: the title bar and the ribbon leave; the strip brings the ribbon back over the slide. */
.pptxv[data-ribbon-display='fullScreen'] > .pptxv-titlebar,
.pptxv[data-ribbon-display='fullScreen'] > .pptxv-qat-row,
.pptxv-ribbon[data-ribbon-display='fullScreen']:not([data-ribbon-peek]) { display: none; }
.pptxv-ribbon[data-ribbon-display='fullScreen'][data-ribbon-peek] {
	position: absolute;
	top: 0;
	left: 0;
	right: 0;
	z-index: 35;
	background: linear-gradient(var(--pptx-card), var(--pptx-card)), var(--pptx-background);
	box-shadow: 0 8px 20px rgb(0 0 0 / 0.2);
}
.pptxv-ribbon-reveal {
	flex: none;
	display: flex;
	align-items: center;
	justify-content: center;
	height: 14px;
	padding: 0;
	border: none;
	border-bottom: 1px solid var(--pptx-border);
	background: var(--pptx-card);
	color: var(--pptx-muted-foreground);
	font: inherit;
	font-size: 14px;
	line-height: 1;
	letter-spacing: 2px;
	cursor: pointer;
}
.pptxv-ribbon-reveal:hover { background: var(--pptx-accent); color: var(--pptx-foreground); }
`;
