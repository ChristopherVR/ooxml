import { translationsEn } from 'ooxml-ui/pptx/i18n';
/**
 * A raw-XML table whose runs name a weight variant as their own family ("DM
 * Sans Light") must render those runs at the family's weight, as the
 * structured-model path does: the XML layer resolves the cell weight from the
 * first run's raw `b` attribute, so a `b="0"` run used to pin the cell to 400.
 */
import type { TablePptxElement, XmlObject } from 'pptx-viewer-core';
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it, vi } from 'vitest';

import { renderTableElement } from './table-render';

vi.mock<typeof import('react-i18next')>(import('react-i18next'), () => ({
	useTranslation: () => ({
		t: (key: string) => translationsEn[key] ?? key,
	}),
}));

const FAMILIES = ['DM Sans Light', 'DM Sans Medium', 'DM Sans Black'];

/** A cell whose run names `typeface`; `b="0"` turns bold off even in a header row. */
function cellXml(text: string, typeface: string, notBold: boolean): XmlObject {
	const rPr = { ...(notBold ? { '@_b': '0' } : {}), 'a:latin': { '@_typeface': typeface } };
	return { 'a:txBody': { 'a:p': { 'a:r': { 'a:rPr': rPr, 'a:t': text } } } } as XmlObject;
}

/** A header row and one body row; every cell mixes the three families. */
function rawXmlTable(families: readonly string[] = FAMILIES): TablePptxElement {
	const rowXml = (prefix: string, notBold: boolean): XmlObject =>
		({
			'a:tc': [
				cellXml(`${prefix}1`, families[0], notBold),
				cellXml(`${prefix}2`, families[0], notBold),
			],
		}) as XmlObject;
	const cell = (prefix: string) => ({
		text: prefix,
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
			rows: [{ cells: [cell('H'), cell('H')] }, { cells: [cell('A'), cell('A')] }],
		},
		rawXml: {
			'a:graphic': {
				'a:graphicData': {
					'a:tbl': {
						'a:tblGrid': { 'a:gridCol': [{ '@_w': '1828800' }, { '@_w': '1828800' }] },
						// The header row leaves bold to the table style.
						'a:tr': [rowXml('H', false), rowXml('A', true)],
					},
				},
			},
		} as XmlObject,
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

describe('raw-XML table weight variants', () => {
	it('renders each regular run at its family weight, and keeps the header bold', () => {
		const markup = renderToStaticMarkup(renderTableElement(rawXmlTable(), { color: '#000000' }));
		const [header, , body] = weights(markup);
		expect(body).toStrictEqual({ cell: '300', runs: ['inherit', '500', '900'] });
		expect(header.cell).toMatch(/^(bold|700)$/u);
		expect(header.runs).toStrictEqual(['inherit', 'inherit', '900']);
	});

	it('does not read a Black first run as a bold cell', () => {
		const families = ['DM Sans Black', 'DM Sans Light', 'DM Sans Medium'];
		const markup = renderToStaticMarkup(
			renderTableElement(rawXmlTable(families), { color: '#000000' }),
		);
		const [, , body] = weights(markup);
		expect(body).toStrictEqual({ cell: '900', runs: ['inherit', '300', '500'] });
	});
});
