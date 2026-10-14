import type { PptxElement, PptxTheme } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { buildTableViewModel, tableStyleContextFor } from './table-renderer-helpers';

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

describe('angular table rendering with theme style references', () => {
	it('builds the style context with the theme format scheme', () => {
		const ctx = tableStyleContextFor(undefined, THEME);
		expect(ctx.formatScheme).toBe(THEME.formatScheme);
		expect(ctx.fontScheme).toBeUndefined();
	});

	it('draws the Themed Style gridlines and background in the cell view-model', () => {
		const rows = buildTableViewModel(TABLE, tableStyleContextFor(undefined, THEME));
		const style = rows[1].cells[1].tdStyle;
		expect(String(style['border-top'])).toContain('#4472C4');
		expect(style['background-color']).toBeDefined();
	});
});
