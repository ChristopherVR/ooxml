/**
 * Layout for the table-cell panels of the inspector and the spacing of the
 * "Theme Colors" grid.
 *
 * `table-cell-formatting.ts`, `table-cell-fill-controls.ts` and
 * `table-structure-controls.ts` build their fields with local helpers that
 * emit bare `<label>`, `<input>` and `<button>` elements. With no rules for
 * their classes the cell panel ran together as one line of inline text. The
 * Theme Colors grid (`theme-color-swatch-grid.ts`) is the shared
 * `pptx-ui-color-grid`, which draws and sizes its own swatches.
 */
export const INSPECTOR_TABLE_CELL_CSS = `
/* ── Table cell panels ───────────────────────────────────────────────── */
.pptxv-inspector-table-cell,
.pptxv-inspector-table-cell-fill {
	display: grid;
	gap: 6px;
	margin-top: 8px;
}
.pptxv-inspector-table-cell[hidden],
.pptxv-inspector-table-cell-fill[hidden],
.pptxv-inspector-table-structure[hidden] { display: none; }
.pptxv-inspector-table-cell > strong { font-size: 11px; font-weight: 600; }
.pptxv-inspector-table-cell > label,
.pptxv-inspector-table-cell-fill > label {
	display: flex;
	align-items: center;
	justify-content: space-between;
	gap: 8px;
	min-width: 0;
	color: var(--pptx-muted-foreground);
	font-size: 11px;
}
.pptxv-inspector-table-cell input[type='number'],
.pptxv-inspector-table-cell-fill input[type='number'],
.pptxv-inspector-table-structure input[type='number'] {
	box-sizing: border-box;
	width: 64px;
	height: 24px;
	padding: 2px 6px;
	border: 1px solid var(--pptx-border);
	border-radius: var(--pptx-radius);
	background: var(--pptx-background);
	color: var(--pptx-foreground);
	font: inherit;
	font-size: 11px;
}
.pptxv-inspector-table-cell-fill > input[type='number'] { justify-self: end; }
.pptxv-inspector-table-cell input[type='color'],
.pptxv-inspector-table-cell-fill input[type='color'] {
	flex: none;
	width: 28px;
	height: 20px;
	padding: 0;
	border: 1px solid var(--pptx-border);
	border-radius: 4px;
	background: none;
	cursor: pointer;
}
.pptxv-inspector-table-structure {
	display: grid;
	grid-template-columns: repeat(3, minmax(0, 1fr));
	gap: 4px;
	margin-top: 8px;
}
.pptxv-inspector-table-structure button {
	min-width: 0;
	padding: 4px 6px;
	border: 1px solid var(--pptx-border);
	border-radius: var(--pptx-radius);
	background: var(--pptx-muted);
	color: inherit;
	font: inherit;
	font-size: 11px;
	cursor: pointer;
}
.pptxv-inspector-table-structure button:hover:not(:disabled) { background: var(--pptx-accent); }
.pptxv-inspector-table-structure button:disabled { opacity: 0.5; cursor: default; }
.pptxv-inspector-table-structure input[type='number'] { width: 100%; }

/* ── Theme Colors grid (ui/theme-color-swatch-grid.ts) ───────────────── */
/* The grid itself is the shared pptx-ui-color-grid; only its spacing is host-owned. */
.pptxv-theme-swatch-grid { margin-top: 4px; }
`;
