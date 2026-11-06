import type { PptxTableData, PptxThemeColorScheme, PptxThemeFormatScheme } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { tableFrameCss } from './table-style';

const THEMED_2_ACCENT_1 = '{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}';
const MEDIUM_2_ACCENT_1 = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';

const colorScheme = { accent1: '#156082', dk1: '#000000', lt1: '#FFFFFF' } as PptxThemeColorScheme;

/** The Office theme's third effect style: an outer shadow of 4.5pt blur, 1.5pt down, 63% black. */
const formatScheme: PptxThemeFormatScheme = {
	fillStyles: [],
	lineStyles: [],
	effectStyles: [
		{ rawNode: {} },
		{ rawNode: {} },
		{
			shadowColor: '#000000',
			shadowOpacity: 0.63,
			shadowBlur: 6,
			shadowOffsetX: 0,
			shadowOffsetY: 2,
			rawNode: {},
		},
	],
	backgroundFillStyles: [],
};

const table = (tableStyleId: string): PptxTableData => ({
	rows: [],
	columnWidths: [],
	tableStyleId,
});

describe('table style drop shadow (a:tblBg/a:effectRef)', () => {
	it('draws Themed Style 2 with the theme effect style shadow on its frame', () => {
		const css = tableFrameCss(table(THEMED_2_ACCENT_1), { colorScheme, formatScheme });
		// 6px DrawingML blur is a 3.9px CSS blur (effect-profile.ts).
		expect(css.boxShadow).toBe('0px 2px 4px rgba(0, 0, 0, 0.63)');
	});

	it('draws nothing for a style without a background effect or without a theme', () => {
		expect(tableFrameCss(table(MEDIUM_2_ACCENT_1), { colorScheme, formatScheme })).toStrictEqual(
			{},
		);
		expect(tableFrameCss(table(THEMED_2_ACCENT_1), { colorScheme })).toStrictEqual({});
	});
});
