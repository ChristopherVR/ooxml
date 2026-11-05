import type { PptxTableCell } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { borderCssWidthPx, insetPaddingForBorders } from './table-cell-border-inset';
import { tableCellCss } from './table-cell-css';

const POSITION = { rowIndex: 0, cellIndex: 0, rowCount: 1, columnCount: 1 };

function cell(style: Record<string, unknown>): PptxTableCell {
	return { text: 'x', style } as unknown as PptxTableCell;
}

describe('table cell border inset', () => {
	it('reads the painted width of a border value', () => {
		expect(borderCssWidthPx('max(1px, var(--pptx-device-px, 0px)) solid #999')).toBe(1);
		expect(borderCssWidthPx('2px solid red')).toBe(2);
		expect(borderCssWidthPx(undefined)).toBe(0);
	});

	it('takes half a border out of the default margins so a row is content + margins', () => {
		const css = tableCellCss(
			undefined,
			cell({ borderTopWidth: 1, borderBottomWidth: 1, borderTopColor: '#999' }),
			POSITION,
		);
		// default marT/marB 4.8px, one px border on top, none declared below
		expect(css.paddingTop).toBe('4.3px');
		expect(css.paddingLeft).toBe('9.6px');
	});

	it('uses the default width of a border that only declares a colour', () => {
		const css = tableCellCss(
			undefined,
			cell({ borderBottomColor: '#999', borderBottomWidth: undefined }),
			POSITION,
		);
		expect(css.paddingBottom).toBe('4.3px');
	});

	it('insets an explicit margin by the border the table style paints', () => {
		const css: Record<string, unknown> = { paddingTop: '4.8px', borderTop: '2px solid red' };
		insetPaddingForBorders(css, { marginTop: 10 } as never);
		expect(css.paddingTop).toBe('9px');
	});
});
