import type { PptxElement, PptxTheme, TablePptxElement } from 'pptx-viewer-core';
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it } from 'vitest';

import { getTableCellBandStyle } from './table-band-style';
import { renderTableFromTableData } from './table-render-data';

const THEMED_1_ACCENT_1 = '{3C2FFA5D-87B4-456A-9821-1D502468CF0F}';

/** A theme whose fill styles are tints of phClr and whose lines are thin phClr. */
const THEME = {
	colorScheme: { dk1: '#000000', lt1: '#FFFFFF', accent1: '#4472C4' },
	formatScheme: {
		fillStyles: [1, 2, 3].map((n) => ({
			kind: 'solid',
			rawNode: { 'a:schemeClr': { '@_val': 'phClr', 'a:tint': { '@_val': String(n * 20000) } } },
		})),
		lineStyles: [{ width: 1, rawNode: { 'a:solidFill': { 'a:schemeClr': { '@_val': 'phClr' } } } }],
		effectStyles: [],
		backgroundFillStyles: [],
	},
} as unknown as PptxTheme;

/** A 3 x 3 Themed Style 1 table with every Table Style Option off. */
const TABLE = {
	id: 't',
	type: 'table',
	x: 0,
	y: 0,
	width: 300,
	height: 90,
	tableData: {
		tableStyleId: THEMED_1_ACCENT_1,
		columnWidths: [1 / 3, 1 / 3, 1 / 3],
		rows: [0, 1, 2].map(() => ({ cells: [{ text: 'a' }, { text: 'b' }, { text: 'c' }] })),
	},
} as unknown as PptxElement;

describe('react table rendering with theme style references', () => {
	it('passes the theme format scheme so Themed Style draws its gridlines', () => {
		const css = getTableCellBandStyle(TABLE, 1, 1, 3, 3, { theme: THEME });
		expect(String(css?.borderTop)).toContain('#4472C4');
		expect(css?.backgroundColor).toBeDefined();
	});

	it('paints the Themed Style table background in the rendered table', () => {
		const html = renderToStaticMarkup(
			<>
				{renderTableFromTableData(TABLE as TablePptxElement, {}, { styleCtx: { theme: THEME } })}
			</>,
		);
		// The table background (tblBg fillRef idx 2) and the accent1 gridlines.
		expect(html).toMatch(/<td[^>]*background-color:/u);
		expect(html).toContain('#4472C4');
	});
});
