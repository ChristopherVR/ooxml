import { translationsEn } from 'ooxml-ui/pptx/i18n';
/**
 * A table loaded from a real .pptx renders from its structured `tableData` through
 * the shared `tableCellCss`, so the table style's own header/band text colour shows
 * through when a cell declares no explicit run colour (the normal case: table styles
 * exist so authors need not hand-colour every header cell). The element's plain
 * default text colour must not clobber it.
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

/** A table with a header row, no explicit per-cell colour anywhere. */
function rawXmlTable(): TablePptxElement {
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
			bandedRows: true,
			rows: [
				{ cells: [{ text: 'H1' }, { text: 'H2' }] },
				{ cells: [{ text: 'A1' }, { text: 'A2' }] },
			],
		},
	} as TablePptxElement;
}

/** Every `<td …style="…">` fragment, in document order. */
function cellStyles(markup: string): string[] {
	return Array.from(markup.matchAll(/<td[^>]*style="([^"]*)"/gu)).map((match) => match[1]);
}

describe('table band colour', () => {
	it("lets the header band's own text colour survive an un-styled cell", () => {
		// A non-trivial element default colour: proves the band colour is not
		// merely surviving because there was nothing to clobber it with.
		const markup = renderToStaticMarkup(renderTableElement(rawXmlTable(), { color: '#000000' }));
		const [header] = cellStyles(markup);
		// The default header emphasis (no table-style GUID) paints white text on
		// a dark fill; the fallback bug painted the element's plain default
		// colour instead, silently losing the band's white.
		expect(header).toContain('color:#ffffff');
	});
});
