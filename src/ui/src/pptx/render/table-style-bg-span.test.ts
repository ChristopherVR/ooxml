import type { ParsedTableStyleEntry, PptxTableData } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { resolveCellBorderCss } from './table-style-borders';
import { spanTableBackground, tableCellSlice } from './table-style-bg-span';

const table = (): PptxTableData =>
	({
		rows: [
			{ height: 10, cells: [] },
			{ height: 20, cells: [] },
			{ height: 10, cells: [] },
		],
		columnWidths: [0.5, 0.5],
	}) as PptxTableData;

describe('table background gradient spans the whole table', () => {
	const gradient = 'linear-gradient(180deg, #111 0%, #eee 100%)';

	it('locates a cell as fractions of the table', () => {
		expect(tableCellSlice(table(), 1, 1)).toEqual({
			left: 0.5,
			top: 0.25,
			width: 0.5,
			height: 0.5,
		});
	});

	it('sizes and positions the gradient to the table, not the cell', () => {
		const css: Record<string, string | number> = { background: gradient };
		spanTableBackground(css, gradient, table(), 1, 0);
		expect(css.backgroundSize).toBe('200% 200%');
		expect(css.backgroundPosition).toBe('0% 50%');
		const first: Record<string, string | number> = { background: gradient };
		spanTableBackground(first, gradient, table(), 0, 0);
		expect(first.backgroundPosition).toBe('0% 0%');
		const last: Record<string, string | number> = { background: gradient };
		spanTableBackground(last, gradient, table(), 2, 1);
		expect(last.backgroundPosition).toBe('100% 100%');
	});

	it('keeps a translucent tint layered above the table gradient at the cell box', () => {
		const css: Record<string, string | number> = {
			background: `linear-gradient(rgba(1, 2, 3, 0.4), rgba(1, 2, 3, 0.4)), ${gradient}`,
		};
		spanTableBackground(css, gradient, table(), 0, 0);
		expect(css.backgroundSize).toBe('auto, 200% 400%');
	});

	it('leaves a cell alone when an opaque fill replaced the gradient', () => {
		const css: Record<string, string | number> = { backgroundColor: '#fff' };
		spanTableBackground(css, gradient, table(), 0, 0);
		expect(css.backgroundSize).toBeUndefined();
	});
});

describe('compound double line (cmpd="dbl")', () => {
	const entry: ParsedTableStyleEntry = {
		styleId: '{X}',
		lastRowBorders: { top: { width: 4, compound: 'dbl', color: '#156082' } },
	};
	const data = { ...table(), lastRow: true } as PptxTableData;

	it('draws a double border where the style asks for one', () => {
		const css = resolveCellBorderCss(
			entry,
			data,
			{ rowIndex: 2, cellIndex: 0, rowCount: 3, columnCount: 2 },
			() => undefined,
		);
		expect(css?.borderTop).toBe('max(4px, var(--pptx-device-px, 0px)) double #156082');
	});
});
