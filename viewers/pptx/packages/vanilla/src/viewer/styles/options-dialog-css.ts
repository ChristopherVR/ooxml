/**
 * Styles for the panes the PowerPoint "File > Options" dialog supplies on top of the shared
 * options dialog (fonts, Customize Ribbon, Quick Access, Add-ins), the title-bar Quick Access strip, the end-of-slide-show
 * black slide, the in-show context menu, and the option-driven viewer root
 * classes (`pptxv-no-hw-accel`, `pptxv-compat-display`, `pptxv-reduced-motion`,
 * `pptxv-no-show-popup`).
 */
export const OPTIONS_DIALOG_CSS = `
/* ── Sections + control rows ── */
.pptxv-options-section { margin: 0 0 16px; }
.pptxv-options-section > h3 { margin: 0 0 4px; padding-bottom: 3px; border-bottom: 1px solid var(--pptx-border); color: var(--pptx-muted-foreground); font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: .04em; }
.pptxv-options-section-desc { margin: 0 0 8px; color: var(--pptx-muted-foreground); font-size: 11px; }
/* File > Options > General > Fonts: families added from a local file. */
.pptxv-options-font-list {
	margin: 4px 0 0;
	padding: 0;
	list-style: none;
	display: flex;
	flex-direction: column;
	gap: 2px;
	font-size: 12px;
}
.pptxv-options-action { padding: 6px 11px; border: 1px solid var(--pptx-border); border-radius: 4px; background: var(--pptx-muted); color: inherit; font: inherit; font-size: 12px; cursor: pointer; }
.pptxv-options-action:hover:not(:disabled) { background: var(--pptx-accent); color: var(--pptx-accent-foreground); }
.pptxv-options-action:disabled { opacity: .45; cursor: default; }

/* ── Customize Ribbon pane ── */
.pptxv-options-ribbon-list { display: flex; flex-direction: column; gap: 2px; padding: 6px; border: 1px solid var(--pptx-border); border-radius: 6px; }
.pptxv-options-ribbon-list .pptxv-parity-check.is-locked { opacity: .55; cursor: not-allowed; }

/* ── Quick Access pane ── */
.pptxv-options-qa { display: flex; align-items: stretch; gap: 10px; }
.pptxv-options-qa-col { flex: 1; min-width: 0; }
.pptxv-options-qa-title { margin: 0 0 4px; color: var(--pptx-muted-foreground); font-size: 11px; font-weight: 600; }
.pptxv-options-qa-list { display: flex; flex-direction: column; gap: 2px; height: 190px; overflow-y: auto; padding: 4px; border: 1px solid var(--pptx-border); border-radius: 6px; }
.pptxv-options-qa-list button { padding: 6px 8px; border: 0; border-radius: 4px; background: transparent; color: var(--pptx-foreground); font: inherit; font-size: 12px; text-align: left; cursor: pointer; }
.pptxv-options-qa-list button:hover { background: var(--pptx-accent); color: var(--pptx-accent-foreground); }
.pptxv-options-qa-list button.is-selected { background: color-mix(in srgb, var(--pptx-primary) 16%, transparent); color: var(--pptx-primary); }
.pptxv-options-qa-arrows { display: flex; flex-direction: column; justify-content: center; gap: 6px; }

/* ── Add-ins pane ── */
.pptxv-options-addins-head { display: grid; grid-template-columns: 1fr auto auto; gap: 8px; padding: 0 8px 4px; border-bottom: 1px solid var(--pptx-border); color: var(--pptx-muted-foreground); font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: .04em; }
.pptxv-options-addins h4 { margin: 12px 0 4px; font-size: 11px; }
.pptxv-options-addins table { width: 100%; border-collapse: collapse; text-align: left; }
.pptxv-options-addins td { padding: 5px 8px; border-bottom: 1px solid color-mix(in srgb, var(--pptx-border) 55%, transparent); font-size: 11px; }
.pptxv-options-addins td:nth-child(2) { color: var(--pptx-muted-foreground); font-family: ui-monospace, monospace; font-size: 10px; }
.pptxv-options-addins td:nth-child(3) { color: var(--pptx-muted-foreground); }
.pptxv-options-addins tr { cursor: pointer; }
.pptxv-options-addins tr:hover { background: var(--pptx-accent); }
.pptxv-options-addins tr.is-selected { background: color-mix(in srgb, var(--pptx-primary) 12%, transparent); }
.pptxv-options-addins-empty { padding: 4px 8px; color: var(--pptx-muted-foreground); font-style: italic; }
.pptxv-options-addins-detail { margin-top: 10px; padding: 10px; border: 1px solid var(--pptx-border); border-radius: 6px; background: var(--pptx-muted); }
.pptxv-options-addins-detail p { margin: 0 0 4px; font-size: 11px; }
.pptxv-options-addins-detail p:last-child { margin: 0; color: var(--pptx-muted-foreground); font-family: ui-monospace, monospace; font-size: 10px; }

/* ── End of slide show (black slide) ── */
.pptxv-endshow { position: absolute; inset: 0; z-index: 60; display: grid; place-items: center; border: 0; background: #000; color: #e5e7eb; font-size: 13px; cursor: pointer; }

/* ── Option-driven viewer root classes ── */
.pptxv.pptxv-reduced-motion *, .pptxv.pptxv-reduced-motion *::before, .pptxv.pptxv-reduced-motion *::after { animation-duration: 0.01ms !important; animation-iteration-count: 1 !important; transition-duration: 0.01ms !important; scroll-behavior: auto !important; }
.pptxv.pptxv-no-hw-accel * { will-change: auto !important; backface-visibility: visible !important; perspective: none !important; }
.pptxv.pptxv-compat-display .pptxv-stage-wrap { box-shadow: none; }
.pptxv.pptxv-compat-display .pptxv-stage { image-rendering: crisp-edges; text-rendering: optimizeSpeed; }
.pptxv.pptxv-no-show-popup .pptxv-presentation-touch-prev,
.pptxv.pptxv-no-show-popup .pptxv-presentation-touch-next,
.pptxv.pptxv-no-show-popup .pptxv-presentation-touch-counter { display: none !important; }
/* The desktop hover toolbar (createPresentationToolbar) sets its own inline
   opacity/pointer-events on mousemove; !important here is required to win
   over that inline style regardless of the JS auto-show timer. */
.pptxv.pptxv-no-show-popup .pptxv-present-toolbar-wrap { opacity: 0 !important; pointer-events: none !important; }
`;
