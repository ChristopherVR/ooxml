import { translationsEn } from 'ooxml-ui/pptx/i18n';
/**
 * A table cell with a painted border takes half of that border out of its default
 * margins (PowerPoint measures a margin from the grid line, so a row is content plus
 * margins tall). React gets it from the shared `tableCellCss`, like the other four
 * bindings, rather than composing cell styles itself.
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

/** One cell whose top and left borders are 2px. */
function borderedTable(): TablePptxElement {
	const style = {
		borderTopWidth: 2,
		borderTopColor: '#000000',
		borderLeftWidth: 2,
		borderLeftColor: '#000000',
	};
	return {
		id: 'tbl-inset',
		type: 'table',
		x: 0,
		y: 0,
		width: 200,
		height: 60,
		tableData: { columnWidths: [1], rows: [{ cells: [{ text: 'x', style }] }] },
	} as unknown as TablePptxElement;
}

describe('table cell padding', () => {
	it('takes half of the painted border out of the default margins', () => {
		const markup = renderToStaticMarkup(renderTableElement(borderedTable(), {}));
		const style = /<td[^>]*style="([^"]*)"/u.exec(markup)?.[1] ?? '';
		expect(style).toContain('padding-top:3.8px');
		expect(style).toContain('padding-left:8.6px');
		expect(style).toContain('padding-bottom:4.8px');
	});
});
