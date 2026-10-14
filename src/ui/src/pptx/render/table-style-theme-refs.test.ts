import type { PptxTableData, PptxThemeColorScheme, PptxThemeFormatScheme } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildRibbonGallery } from './ribbon-galleries/gallery-registry';
import { getCellDiagonalBorders, getTableCellBandStyle } from './table-style';
import { getBuiltinTableStyle } from './table-style-builtins';
import { applyTableStyleFieldEdit } from './table-style-editor-edit';
import { resolveTableStyleThemeRefs } from './table-style-theme-refs';

const THEMED_1_ACCENT_1 = '{3C2FFA5D-87B4-456A-9821-1D502468CF0F}';
const THEMED_2_ACCENT_3 = '{306799F8-075E-4A3A-A7F6-7FBC6576F1A4}';
const NO_STYLE_TABLE_GRID = '{5940675A-B579-460E-94D1-54222C63F5DA}';

const TEST_COLOR_SCHEME = {
	dk1: '#000000',
	lt1: '#FFFFFF',
	dk2: '#44546A',
	lt2: '#E7E6E6',
	accent1: '#4472C4',
	accent2: '#ED7D31',
	accent3: '#A5A5A5',
	accent4: '#FFC000',
	accent5: '#5B9BD5',
	accent6: '#70AD47',
	hlink: '#0563C1',
	folHlink: '#954F72',
} as PptxThemeColorScheme;

/** A theme whose fill styles are tints of phClr and whose lines are thin phClr. */
const TEST_FORMAT_SCHEME: PptxThemeFormatScheme = {
	fillStyles: [1, 2, 3].map((n) => ({
		kind: 'solid' as const,
		rawNode: { 'a:schemeClr': { '@_val': 'phClr', 'a:tint': { '@_val': String(n * 20000) } } },
	})),
	lineStyles: [1, 2, 3].map((n) => ({
		width: n,
		rawNode: { '@_w': String(n * 9525), 'a:solidFill': { 'a:schemeClr': { '@_val': 'phClr' } } },
	})),
	effectStyles: [],
	backgroundFillStyles: [],
};

function table(styleId: string, flags: Partial<PptxTableData> = {}): PptxTableData {
	return { rows: [], columnWidths: [], tableStyleId: styleId, ...flags };
}

const ctx = { colorScheme: TEST_COLOR_SCHEME, formatScheme: TEST_FORMAT_SCHEME };

describe('table style theme references', () => {
	it('carries the Themed Style tblBg and line references in the built-in catalogue', () => {
		const entry = getBuiltinTableStyle(THEMED_1_ACCENT_1);
		expect(entry?.tableBackground?.fillRef).toStrictEqual({
			idx: 2,
			color: { schemeColor: 'accent1' },
		});
		expect(entry?.wholeTblBorders?.insideH?.lnRef?.idx).toBe(1);
		expect(getBuiltinTableStyle(THEMED_2_ACCENT_3)?.tableBackground?.fillRef?.color).toStrictEqual({
			schemeColor: 'accent3',
		});
	});

	it('paints Themed Style 1 with every option off: background and gridlines', () => {
		const css = getTableCellBandStyle(table(THEMED_1_ACCENT_1), 1, 1, 3, 3, ctx);
		// fillStyleLst[1] = phClr at tint 40%, phClr = accent1.
		expect(css?.backgroundColor).toBeDefined();
		expect(css?.backgroundColor).not.toBe('transparent');
		for (const edge of ['borderTop', 'borderBottom', 'borderLeft', 'borderRight']) {
			expect(String(css?.[edge])).toContain('#4472C4');
		}
	});

	it('keeps drawing nothing for the references without a format scheme', () => {
		const css = getTableCellBandStyle(table(THEMED_1_ACCENT_1), 1, 1, 3, 3, {
			colorScheme: TEST_COLOR_SCHEME,
		});
		expect(css?.borderTop).toBeUndefined();
		expect(css?.backgroundColor ?? 'transparent').toBe('transparent');
		expect(
			getCellDiagonalBorders(undefined, table(THEMED_1_ACCENT_1), {
				rowIndex: 0,
				cellIndex: 0,
				rowCount: 1,
				columnCount: 1,
			}),
		).toBeNull();
	});

	it('returns an entry without references as is, and resolves one with them', () => {
		const plain = getBuiltinTableStyle(NO_STYLE_TABLE_GRID);
		expect(resolveTableStyleThemeRefs(plain, TEST_FORMAT_SCHEME, TEST_COLOR_SCHEME)).toBe(plain);
		const themed = resolveTableStyleThemeRefs(
			getBuiltinTableStyle(THEMED_1_ACCENT_1),
			TEST_FORMAT_SCHEME,
			TEST_COLOR_SCHEME,
		);
		expect(themed?.tableBackground?.fill?.color).toBeDefined();
		expect(themed?.wholeTblBorders?.top).toStrictEqual({ width: 1, color: '#4472C4' });
		expect(
			resolveTableStyleThemeRefs(undefined, TEST_FORMAT_SCHEME, TEST_COLOR_SCHEME),
		).toBeUndefined();
	});

	// The style editor edits the entry in place: a table painted before the
	// edit must repaint with it, with or without theme references.
	it.each([
		['No Style, Table Grid', NO_STYLE_TABLE_GRID],
		['Themed Style 1 - Accent 1', THEMED_1_ACCENT_1],
	])('repaints %s after an in-place style edit', (_name, styleId) => {
		const builtin = getBuiltinTableStyle(styleId);
		if (!builtin) {
			throw new Error(`no built-in style ${styleId}`);
		}
		const entry = structuredClone(builtin);
		const styleCtx = { ...ctx, tableStyleMap: { [entry.styleId]: entry } };
		const paint = () =>
			getTableCellBandStyle(table(entry.styleId), 1, 1, 3, 3, styleCtx)?.backgroundColor;
		expect(paint()).not.toBe('#FF00FF');
		applyTableStyleFieldEdit(entry, 'wholeTbl', {
			kind: 'fillColor',
			hex: '#FF00FF',
			ref: undefined,
		});
		expect(paint()).toBe('#FF00FF');
	});

	it('paints the Themed Style gallery tiles with the theme format scheme', () => {
		const tile = (formatScheme?: PptxThemeFormatScheme) =>
			buildRibbonGallery('tableStyles', {
				element: {
					id: 't',
					type: 'table',
					x: 0,
					y: 0,
					width: 10,
					height: 10,
					tableData: table(THEMED_1_ACCENT_1, { firstRowHeader: true, bandedRows: true }),
				} as never,
				theme: { colorScheme: TEST_COLOR_SCHEME, ...(formatScheme ? { formatScheme } : {}) },
			})
				.sections.flatMap((section) => section.items)
				.find((item) => item.id === THEMED_1_ACCENT_1)?.previewSvg ?? '';
		// A body cell's colour is the tblBg fill style resolved for accent1.
		const body = getTableCellBandStyle(table(THEMED_1_ACCENT_1), 2, 2, 5, 5, ctx);
		const fill = `fill="${String(body?.backgroundColor)}"`;
		expect(tile(TEST_FORMAT_SCHEME)).toContain(fill);
		expect(tile()).not.toContain(fill);
	});
});
