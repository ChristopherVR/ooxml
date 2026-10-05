/**
 * A subscript run and a bordered cell margin, rendered through the full
 * table path so the shared run style and cell padding reach the markup.
 */
import type { PptxTableCell, TablePptxElement } from 'pptx-viewer-core';
import { translationsEn } from 'pptx-viewer-shared/i18n';
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it, vi } from 'vitest';

import { renderTableElement } from './table-render';

vi.mock<typeof import('react-i18next')>(import('react-i18next'), () => ({
	useTranslation: () => ({
		t: (key: string) => translationsEn[key] ?? key,
	}),
}));

function tableWith(cell: PptxTableCell): TablePptxElement {
	return {
		id: 'tbl-text',
		type: 'table',
		x: 0,
		y: 0,
		width: 400,
		height: 200,
		tableData: { columnWidths: [1], rows: [{ cells: [cell] }] },
	} as TablePptxElement;
}

function firstCellStyle(markup: string): string {
	return Array.from(markup.matchAll(/<td[^>]*style="([^"]*)"/gu))[0]?.[1] ?? '';
}

describe('table cell text', () => {
	it('lowers and shrinks a subscript run', () => {
		const sub = { text: '2', fontSize: 10, baseline: -25000 };
		const markup = renderToStaticMarkup(
			renderTableElement(
				tableWith({ text: 'CO2', textRuns: [{ text: 'CO', fontSize: 10 }, sub] }),
				{},
			),
		);
		expect(markup).toMatch(
			/<span style="[^"]*font-size:6.5pt[^"]*vertical-align:sub[^"]*">2<\/span>/u,
		);
	});

	it('takes half of the border width out of the cell padding', () => {
		const markup = renderToStaticMarkup(
			renderTableElement(
				tableWith({ text: '42', style: { marginRight: 6, borderRightWidth: 2 } }),
				{},
			),
		);
		expect(firstCellStyle(markup)).toContain('padding-right:5px');
	});
});
