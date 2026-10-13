/**
 * Layout for the table-cell panels of the inspector and for the "Theme
 * Colors" swatch grid.
 *
 * `table-cell-formatting.ts`, `table-cell-fill-controls.ts` and
 * `table-structure-controls.ts` build their fields with local helpers that
 * emit bare `<label>`, `<input>` and `<button>` elements, and
 * `theme-color-swatch-grid.ts` emits empty `<button>`s coloured by an inline
 * background. With no rules for their classes the cell panel ran together as
 * one line of inline text and every theme swatch collapsed to a zero-sized
 * button, so the grid showed only its caption. The sizes match the other
 * bindings' grids (16px swatches, 2px gaps).
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
.pptxv-theme-swatch-grid { display: flex; flex-direction: column; gap: 4px; margin-top: 4px; }
.pptxv-theme-swatch-grid[hidden] { display: none; }
.pptxv-theme-swatch-grid-label { color: var(--pptx-muted-foreground); font-size: 10px; }
.pptxv-theme-swatch-grid-rows { display: flex; flex-direction: column; gap: 2px; }
.pptxv-theme-swatch-grid-row { display: flex; gap: 2px; }
.pptxv-theme-swatch-grid-swatch,
.pptxv-theme-swatch-grid-empty { flex: none; width: 16px; height: 16px; }
.pptxv-theme-swatch-grid-swatch {
	box-sizing: border-box;
	padding: 0;
	border: 1px solid var(--pptx-border);
	border-radius: 2px;
	cursor: pointer;
}
.pptxv-theme-swatch-grid-swatch:hover:not(:disabled) { transform: scale(1.1); }
.pptxv-theme-swatch-grid-swatch:disabled { opacity: 0.4; cursor: not-allowed; }
.pptxv-theme-swatch-grid-swatch.is-selected { outline: 2px solid var(--pptx-ring); outline-offset: 1px; }
`;
