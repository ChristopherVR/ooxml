import { translationsEn } from 'ooxml-ui/pptx/i18n';
/**
 * A raw-XML table cell with a painted border takes half of that border out of
 * its default margins, as `tableCellCss` does on the structured-model path
 * (PowerPoint measures a margin from the grid line, so a row is content plus
 * margins tall). The other four bindings get it from the shared `tableCellCss`;
 * React composes raw-XML cells itself.
 */
import type { TablePptxElement, XmlObject } from 'pptx-viewer-core';
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it, vi } from 'vitest';

import { renderTableElement } from './table-render';

vi.mock<typeof import('../i18n/use-translation')>(import('../i18n/use-translation'), () => ({
	useTranslation: () => ({
		t: (key: string) => translationsEn[key] ?? key,
	}),
}));

/** One cell whose top border is 2px (19050 EMU at 9525 EMU per px). */
function borderedTable(): TablePptxElement {
	const line = (tag: string) => ({
		[tag]: { '@_w': '19050', 'a:solidFill': { 'a:srgbClr': { '@_val': '000000' } } },
	});
	const tc = {
		'a:txBody': { 'a:p': { 'a:r': { 'a:rPr': {}, 'a:t': 'x' } } },
		'a:tcPr': { ...line('a:lnT'), ...line('a:lnL') },
	} as XmlObject;
	return {
		id: 'tbl-inset',
		type: 'table',
		x: 0,
		y: 0,
		width: 200,
		height: 60,
		tableData: { columnWidths: [1], rows: [{ cells: [{ text: 'x' }] }] },
		rawXml: {
			'a:graphic': {
				'a:graphicData': {
					'a:tbl': {
						'a:tblGrid': { 'a:gridCol': [{ '@_w': '1828800' }] },
						'a:tr': [{ '@_h': '371475', 'a:tc': [tc] }],
					},
				},
			},
		} as XmlObject,
	} as TablePptxElement;
}

describe('raw-XML table cell padding', () => {
	it('takes half of the painted border out of the default margins', () => {
		const markup = renderToStaticMarkup(renderTableElement(borderedTable(), {}));
		const style = /<td[^>]*style="([^"]*)"/u.exec(markup)?.[1] ?? '';
		expect(style).toContain('padding-top:3.8px');
		expect(style).toContain('padding-left:8.6px');
		expect(style).toContain('padding-bottom:4.8px');
	});
});
