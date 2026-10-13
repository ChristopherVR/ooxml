import { getTableCellBandStyle } from 'ooxml-ui/pptx';
import type { PptxTableData, PptxThemeColorScheme, PptxThemeFormatScheme } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { renderContextSourceFor, tableStyleContextFrom } from './render-context';

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

describe('svelte table style context with theme style references', () => {
	it('carries the format scheme so Themed Style draws its gridlines', () => {
		const ctx = tableStyleContextFrom({
			getColorScheme: () => COLOR_SCHEME,
			getTableStyleMap: () => undefined,
			getFormatScheme: () => FORMAT_SCHEME,
		});
		expect(ctx?.formatScheme).toBe(FORMAT_SCHEME);
		const table: PptxTableData = { rows: [], columnWidths: [], tableStyleId: THEMED_1_ACCENT_1 };
		const css = getTableCellBandStyle(table, 1, 1, 3, 3, ctx);
		expect(String(css?.borderTop)).toContain('#4472C4');
	});

	it('reads the format scheme from the loaded theme', () => {
		const loader = {
			colorScheme: COLOR_SCHEME,
			tableStyleMap: undefined,
			presentationTheme: { colorScheme: COLOR_SCHEME, formatScheme: FORMAT_SCHEME },
			canvasSize: undefined,
		};
		expect(tableStyleContextFrom(renderContextSourceFor(loader))?.formatScheme).toBe(FORMAT_SCHEME);
	});

	it('is undefined without a source', () => {
		expect(tableStyleContextFrom(undefined)).toBeUndefined();
	});
});
