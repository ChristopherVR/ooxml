import type { PptxElement, PptxTableData } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildRibbonGallery } from './gallery-registry';
import { OFFICE_THEME_COLORS } from './gallery-theme';
import { tableStyleTileOptions, tileBorderStroke } from './table-style-tile-svg';

const MEDIUM_2_ACCENT_1 = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';
const TABLE_GRID = '{5940675A-B579-460E-94D1-54222C63F5DA}';
/** Office accent1, the solid header / first-column fill of Medium Style 2 - Accent 1. */
const ACCENT = OFFICE_THEME_COLORS.accent1;

function table(options: Partial<PptxTableData>): PptxElement {
	const tableData: PptxTableData = {
		columnWidths: [1],
		rows: [{ cells: [{ text: 'A' }] }],
		tableStyleId: MEDIUM_2_ACCENT_1,
		...options,
	};
	return { id: 't', type: 'table', x: 0, y: 0, width: 10, height: 10, tableData } as PptxElement;
}

function previewOf(element: PptxElement | null, styleId = MEDIUM_2_ACCENT_1): string {
	const descriptor = buildRibbonGallery('tableStyles', { element });
	const item = descriptor.sections.flatMap((s) => s.items).find((i) => i.id === styleId);
	if (!item) {
		throw new Error(`expected the ${styleId} tile`);
	}
	return item.previewSvg;
}

const preview = (options: Partial<PptxTableData>): string => previewOf(table(options));

/**
 * The fill colours of the 25 cell rectangles, row by row. The tile is 60 x 44
 * with a 2px inset, so each of its 5 x 5 cells is 11.2 x 8.
 */
function cellFills(svg: string): string[] {
	return [...svg.matchAll(/<rect [^>]*width="11.2" height="8" fill="([^"]+)"/gu)].map((m) => m[1]);
}

/** The fill of the cell at `row`, `col`. */
const at = (fills: string[], row: number, col: number): string | undefined => fills[row * 5 + col];

describe('table style gallery previews follow the table options', () => {
	it('draws no header or banding for a table with every option off', () => {
		const fills = cellFills(preview({}));
		expect(fills).toHaveLength(25);
		expect(new Set(fills).size).toBe(1);
		expect(fills).not.toContain(ACCENT);
	});

	it('draws the header row and the total row only when the table has them', () => {
		const header = cellFills(preview({ firstRowHeader: true }));
		expect(header.slice(0, 5)).toStrictEqual(Array(5).fill(ACCENT));
		expect(new Set(header.slice(5)).size).toBe(1);

		const total = cellFills(preview({ lastRow: true }));
		expect(total.slice(20)).toStrictEqual(Array(5).fill(ACCENT));
		expect(new Set(total.slice(0, 20)).size).toBe(1);
	});

	it('draws banded rows and banded columns only when the table has them', () => {
		const rows = cellFills(preview({ bandedRows: true }));
		expect(at(rows, 0, 0)).not.toBe(at(rows, 1, 0));
		expect(at(rows, 0, 0)).toBe(at(rows, 2, 0));
		expect(at(rows, 0, 0)).toBe(at(rows, 0, 1));

		const cols = cellFills(preview({ bandedColumns: true }));
		expect(at(cols, 0, 0)).not.toBe(at(cols, 0, 1));
		expect(at(cols, 0, 0)).toBe(at(cols, 0, 2));
		expect(at(cols, 0, 0)).toBe(at(cols, 1, 0));
	});

	it('draws the first and last column only when the table has them', () => {
		const first = cellFills(preview({ firstCol: true }));
		const last = cellFills(preview({ lastCol: true }));
		for (let row = 0; row < 5; row++) {
			expect(at(first, row, 0)).toBe(ACCENT);
			expect(at(first, row, 1)).not.toBe(ACCENT);
			expect(at(last, row, 4)).toBe(ACCENT);
			expect(at(last, row, 3)).not.toBe(ACCENT);
		}
	});

	it('shows the header row and banded rows default without a table', () => {
		expect(previewOf(null)).toBe(preview({ firstRowHeader: true, bandedRows: true }));
	});

	it('reads every option of the table as a boolean', () => {
		expect(tableStyleTileOptions({ rows: [], columnWidths: [], lastCol: true })).toStrictEqual({
			firstRowHeader: false,
			lastRow: false,
			firstCol: false,
			lastCol: true,
			bandedRows: false,
			bandedColumns: false,
		});
	});

	it('draws the style gridlines', () => {
		// "No Style, Table Grid" is only lines; the renderer writes each one as
		// `max(<w>px, var(--...)) solid <colour>`.
		const grid = previewOf(table({}), TABLE_GRID);
		expect((grid.match(/<line /gu) ?? []).length).toBeGreaterThan(0);
	});

	it('reads the screenBorder form and skips zero-width and missing lines', () => {
		expect(tileBorderStroke('max(1px, var(--pptx-device-px, 0px)) solid #FFFFFF')).toStrictEqual({
			width: 0.5,
			color: '#FFFFFF',
		});
		expect(tileBorderStroke('4px solid #000000')).toStrictEqual({ width: 2, color: '#000000' });
		expect(tileBorderStroke('0px solid #000000')).toBeNull();
		expect(tileBorderStroke('none')).toBeNull();
		expect(tileBorderStroke(undefined)).toBeNull();
	});
});
