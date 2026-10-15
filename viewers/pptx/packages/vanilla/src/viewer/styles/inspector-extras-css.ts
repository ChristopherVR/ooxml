/**
 * Layout for the inspector parts whose classes had no rule of their own: the
 * media preview, poster and metadata, the animation sound and after-animation
 * rows, the slide-size rescale prompt, the pattern fill panel, the alt-text
 * field, the tag toggle label, the SmartArt node hierarchy buttons and the
 * nested-deck tab of the OLE editor dialog (whose rule named a `-deck-row`
 * class the dialog never renders).
 *
 * Emitted after the other inspector sheets; each rule is a plain class so it
 * wins over the zero-specificity baseline in `inspector-base-css.ts`.
 */
export const INSPECTOR_EXTRAS_CSS = `
/* ── Media ───────────────────────────────────────────────────────────── */
.pptxv-media-preview {
	display: block;
	width: 100%;
	max-height: 160px;
	margin-bottom: 8px;
	border: 1px solid var(--pptx-border);
	border-radius: var(--pptx-radius);
	background: #000;
}
.pptxv-media-preview[hidden] { display: none; }
.pptxv-media-poster {
	display: block;
	max-width: 100%;
	max-height: 96px;
	margin-bottom: 8px;
	border-radius: var(--pptx-radius);
	object-fit: contain;
}
.pptxv-media-poster[hidden],
.pptxv-media-poster:not([src]) { display: none; }
.pptxv-media-reset-trim {
	align-self: flex-start;
	margin: 2px 0 8px;
	padding: 2px 8px;
	border: 1px solid var(--pptx-border);
	border-radius: var(--pptx-radius);
	background: var(--pptx-secondary);
	color: var(--pptx-secondary-foreground);
	font: inherit;
	font-size: 11px;
	cursor: pointer;
}
.pptxv-media-reset-trim:disabled { opacity: 0.5; cursor: default; }
.pptxv-media-metadata {
	display: block;
	margin-top: 6px;
	color: var(--pptx-muted-foreground);
	font-size: 10px;
	font-variant-numeric: tabular-nums;
	white-space: pre-line;
}

/* ── Animation rows ──────────────────────────────────────────────────── */
.pptxv-effect-sound-row,
.pptxv-after-animation-row {
	display: flex;
	flex-direction: column;
	gap: 4px;
	margin-bottom: 8px;
	color: var(--pptx-muted-foreground);
	font-size: 11px;
}
.pptxv-effect-sound-controls {
	display: flex;
	align-items: center;
	gap: 6px;
	min-width: 0;
}
.pptxv-effect-sound-controls > :first-child { flex: 1; min-width: 0; }
.pptxv-effect-sound-preview {
	flex: none;
	width: 26px;
	height: 26px;
	padding: 0;
	border: 1px solid var(--pptx-border);
	border-radius: var(--pptx-radius);
	background: var(--pptx-secondary);
	color: var(--pptx-secondary-foreground);
	font-size: 10px;
	cursor: pointer;
}
.pptxv-effect-sound-preview:disabled { opacity: 0.5; cursor: default; }
/* The picker opens from the "Other sound..." option; the input itself stays out of sight. */
.pptxv-effect-sound-file-input {
	position: absolute;
	width: 1px;
	height: 1px;
	overflow: hidden;
	clip-path: inset(50%);
	white-space: nowrap;
}
.pptxv-after-animation-row > label,
.pptxv-after-animation-color {
	display: flex;
	align-items: center;
	justify-content: space-between;
	gap: 8px;
}
.pptxv-after-animation-row > label > :last-child { flex: 0 1 60%; min-width: 0; }
.pptxv-after-animation-color[hidden] { display: none; }

/* ── Slide size: rescale prompt ──────────────────────────────────────── */
.pptxv-slide-size-rescale {
	display: grid;
	gap: 6px;
	margin-top: 8px;
	padding: 8px;
	border: 1px solid var(--pptx-primary);
	border-radius: 6px;
	background: color-mix(in srgb, var(--pptx-primary) 8%, transparent);
}
.pptxv-slide-size-rescale[hidden] { display: none; }
.pptxv-slide-size-rescale-title { color: var(--pptx-foreground); font-size: 12px; font-weight: 600; }
.pptxv-slide-size-rescale-description { margin: 0; color: var(--pptx-muted-foreground); font-size: 11px; line-height: 1.4; }
.pptxv-slide-size-rescale-buttons { display: grid; grid-template-columns: repeat(auto-fit, minmax(96px, 1fr)); gap: 6px; }
.pptxv-slide-size-rescale-btn {
	min-height: 28px;
	padding: 4px 8px;
	border: 1px solid var(--pptx-border);
	border-radius: var(--pptx-radius);
	background: var(--pptx-background);
	color: var(--pptx-foreground);
	font: inherit;
	font-size: 11px;
	cursor: pointer;
}
.pptxv-slide-size-rescale-btn:hover { border-color: var(--pptx-primary); background: var(--pptx-accent); }

/* ── Fill: pattern panel, alt text, tags ─────────────────────────────── */
.pptxv-inspector-pattern {
	display: grid;
	gap: 6px;
	margin: 4px 0 8px;
	padding-left: 4px;
	border-left: 2px solid var(--pptx-border);
}
.pptxv-inspector-pattern[hidden],
.pptxv-inspector-pattern:empty { display: none; }
.pptxv-accessibility-alt {
	display: flex;
	flex-direction: column;
	align-items: stretch;
	gap: 4px;
	margin-bottom: 8px;
}
.pptxv-accessibility-alt-input { width: 100%; }
.pptxv-accessibility-title-input { flex: 1 1 auto; min-width: 0; }
.pptxv-tags-toggle-label {
	color: var(--pptx-foreground);
	font-size: 11px;
	font-weight: 600;
}

/* ── SmartArt text pane: per-node hierarchy actions ──────────────────── */
.pptxv-smartart-hierarchy {
	display: flex;
	flex-wrap: wrap;
	grid-column: 2 / -1;
	gap: 4px;
	margin-top: 2px;
}
.pptxv-smartart-hierarchy button {
	padding: 1px 6px;
	border: 1px solid var(--pptx-border);
	border-radius: var(--pptx-radius);
	background: transparent;
	color: var(--pptx-muted-foreground);
	font: inherit;
	font-size: 10px;
	cursor: pointer;
}
.pptxv-smartart-hierarchy button:hover { background: var(--pptx-accent); color: var(--pptx-accent-foreground); }

/* ── OLE editor dialog: nested deck tab ──────────────────────────────── */
.pptxv-ole-edit-content { display: flex; flex-direction: column; gap: 8px; min-width: 0; }
.pptxv-ole-edit-deck-slide { display: flex; flex-direction: column; gap: 4px; color: var(--pptx-muted-foreground); }
.pptxv-ole-edit-deck-caption { font-size: 11px; font-weight: 600; }
.pptxv-ole-edit-deck-slide input {
	padding: 6px 8px;
	border: 1px solid var(--pptx-border);
	border-radius: 4px;
	background: var(--pptx-muted);
	color: var(--pptx-foreground);
	font: inherit;
}
`;
