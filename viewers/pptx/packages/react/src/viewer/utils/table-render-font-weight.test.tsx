import { translationsEn } from 'ooxml-ui/pptx/i18n';
/**
 * A table whose cell and runs name a weight variant as their own family ("DM Sans
 * Light") renders them at the family's weight, as shape text does: the shared
 * `tableCellCss` resolves the cell weight once family and boldness are both known.
 */
import type { TablePptxElement } from 'pptx-viewer-core';
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it, vi } from 'vitest';

import { renderTableElement } from './table-render';

vi.mock<typeof import('../i18n/use-translation')>(import('../i18n/use-translation'), () => ({
	useTranslation: () => ({
		t: (key: string) => translationsEn[key] ?? key,
	}),
}));

const FAMILIES = ['DM Sans Light', 'DM Sans Medium', 'DM Sans Black'];

/** A header row and one body row; every cell mixes the three families. */
function weightTable(families: readonly string[] = FAMILIES): TablePptxElement {
	const cell = (prefix: string, notBold: boolean) => ({
		text: prefix,
		// An explicit `b="0"` turns bold off even in a header row.
		style: { fontFamily: families[0], ...(notBold ? { bold: false } : {}) },
		textRuns: families.map((fontFamily, i) => ({ text: `${prefix}${i}`, fontFamily, bold: false })),
	});
	return {
		id: 'tbl-1',
		type: 'table',
		x: 0,
		y: 0,
		width: 400,
		height: 120,
		tableData: {
			columnWidths: [0.5, 0.5],
			firstRowHeader: true,
			rows: [
				{ cells: [cell('H', false), cell('H', false)] },
				{ cells: [cell('A', true), cell('A', true)] },
			],
		},
	} as TablePptxElement;
}

/** `font-weight` of every `<td>` and of every run `<span>` inside it, per row. */
function weights(markup: string): Array<{ cell: string; runs: string[] }> {
	return Array.from(markup.matchAll(/<td[^>]*style="([^"]*)"[^>]*>(.*?)<\/td>/gu)).map((match) => ({
		cell: /font-weight:(\w+)/u.exec(match[1])?.[1] ?? '',
		runs: Array.from(match[2].matchAll(/<span style="([^"]*)"/gu)).map(
			(span) => /font-weight:(\w+)/u.exec(span[1])?.[1] ?? 'inherit',
		),
	}));
}

describe('table weight variants', () => {
	it('renders each regular run at its family weight, and keeps the header bold', () => {
		const markup = renderToStaticMarkup(renderTableElement(weightTable(), { color: '#000000' }));
		const [header, , body] = weights(markup);
		expect(body).toStrictEqual({ cell: '300', runs: ['inherit', '500', '900'] });
		expect(header.cell).toMatch(/^(bold|700)$/u);
		expect(header.runs).toStrictEqual(['inherit', 'inherit', '900']);
	});

	it('does not read a Black first run as a bold cell', () => {
		const families = ['DM Sans Black', 'DM Sans Light', 'DM Sans Medium'];
		const markup = renderToStaticMarkup(
			renderTableElement(weightTable(families), { color: '#000000' }),
		);
		const [, , body] = weights(markup);
		expect(body).toStrictEqual({ cell: '900', runs: ['inherit', '300', '500'] });
	});
});
