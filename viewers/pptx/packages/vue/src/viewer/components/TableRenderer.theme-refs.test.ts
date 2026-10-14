import { mount } from '@vue/test-utils';
import type { PptxElement, PptxThemeColorScheme, PptxThemeFormatScheme } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { TableThemeKey, tableThemeContextFor } from '../composables/table-theme';
import TableRenderer from './TableRenderer.vue';

const THEMED_1_ACCENT_1 = '{3C2FFA5D-87B4-456A-9821-1D502468CF0F}';

const COLOR_SCHEME = { dk1: '#000000', lt1: '#FFFFFF', accent1: '#4472C4' } as PptxThemeColorScheme;

/** A theme whose fill styles are tints of phClr and whose lines are thin phClr. */
const FORMAT_SCHEME = {
	fillStyles: [1, 2, 3].map((n) => ({
		kind: 'solid',
		rawNode: { 'a:schemeClr': { '@_val': 'phClr', 'a:tint': { '@_val': String(n * 20000) } } },
	})),
	lineStyles: [{ width: 1, rawNode: { 'a:solidFill': { 'a:schemeClr': { '@_val': 'phClr' } } } }],
	effectStyles: [],
	backgroundFillStyles: [],
} as unknown as PptxThemeFormatScheme;

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

describe('vue table rendering with theme style references', () => {
	it('provides the theme format scheme from the viewer root', () => {
		const ctx = tableThemeContextFor(
			{ colorScheme: COLOR_SCHEME, formatScheme: FORMAT_SCHEME },
			undefined,
		);
		expect(ctx).toStrictEqual({
			colorScheme: COLOR_SCHEME,
			tableStyleMap: undefined,
			fontScheme: undefined,
			formatScheme: FORMAT_SCHEME,
		});
	});

	it('resolves Themed Style through the injected format scheme', () => {
		const wrapper = mount(TableRenderer, {
			props: { element: TABLE, zIndex: 0 },
			global: {
				provide: {
					[TableThemeKey as symbol]: () => ({
						colorScheme: COLOR_SCHEME,
						formatScheme: FORMAT_SCHEME,
					}),
				},
			},
		});
		const style = wrapper.findAll('tr')[1].findAll('td')[1].attributes('style') ?? '';
		// The table background: tblBg fillRef idx 2, accent1 at a 40% tint.
		expect(style).toMatch(/background-color:/u);
		expect(style).not.toMatch(/background-color: transparent/u);
	});
});
