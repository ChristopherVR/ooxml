import { describe, expect, it } from 'vitest';
import { defaultPdfChoice, pdfOptionsFromChoice, type PdfOptionsChoice } from './pdf-options-form';

const choice = (patch: Partial<PdfOptionsChoice> = {}): PdfOptionsChoice => ({
	...defaultPdfChoice(),
	...patch,
});
const printLayout = { pages: { current: 3, total: 5 }, hasSelection: true };
const draft = { hasSelection: false };

describe('PDF Options step mapping', () => {
	it('maps the default choices to what exportPdf() writes without options', () => {
		expect(pdfOptionsFromChoice(choice(), draft)).toEqual({
			pages: 'all',
			markup: 'none',
			bookmarks: 'headings',
			documentProperties: true,
			taggedPdf: true,
			pdfA: false,
		});
	});

	it('maps publish what, bookmarks, properties, tags and PDF/A', () => {
		expect(
			pdfOptionsFromChoice(
				choice({
					markup: true,
					bookmarkSource: 'wordBookmarks',
					documentProperties: false,
					taggedPdf: false,
					pdfA: true,
				}),
				draft,
			),
		).toMatchObject({
			markup: 'all',
			markupBalloons: 'commentsAndFormatting',
			bookmarks: 'wordBookmarks',
			documentProperties: false,
			taggedPdf: false,
			pdfA: true,
		});
		expect(
			pdfOptionsFromChoice(
				choice({ createBookmarks: false, bookmarkSource: 'wordBookmarks' }),
				draft,
			),
		).toMatchObject({ bookmarks: 'none' });
		// Balloons only apply to the markup; the Document leaves them out.
		expect(pdfOptionsFromChoice(choice({ markup: true, balloons: 'all' }), draft)).toMatchObject({
			markupBalloons: 'all',
		});
		expect(pdfOptionsFromChoice(choice({ balloons: 'all' }), draft)).not.toHaveProperty(
			'markupBalloons',
		);
	});

	it('maps the current page and the selection when the editor can offer them', () => {
		expect(pdfOptionsFromChoice(choice({ range: 'current' }), printLayout)).toMatchObject({
			pages: 'current',
		});
		expect(pdfOptionsFromChoice(choice({ range: 'selection' }), printLayout)).toMatchObject({
			pages: 'selection',
		});
		// Draft view has no current page and nothing is selected: the whole document.
		expect(pdfOptionsFromChoice(choice({ range: 'current' }), draft)).toMatchObject({
			pages: 'all',
		});
		expect(pdfOptionsFromChoice(choice({ range: 'selection' }), draft)).toMatchObject({
			pages: 'all',
		});
	});

	it('maps a page range and points at the field to correct', () => {
		const pages = (from: string, to: string, context: typeof draft = draft) =>
			pdfOptionsFromChoice(choice({ range: 'pages', from, to }), context);
		expect(pages('2', ' 4 ')).toMatchObject({ pages: { from: 2, to: 4 } });
		expect(pages('2', '9', printLayout)).toMatchObject({ pages: { from: 2, to: 5 } });
		expect(pages('0', '2')).toEqual({ invalid: 'from' });
		expect(pages('1.5', '2')).toEqual({ invalid: 'from' });
		expect(pages('', '2')).toEqual({ invalid: 'from' });
		expect(pages('6', '6', printLayout)).toEqual({ invalid: 'from' });
		expect(pages('3', '2')).toEqual({ invalid: 'to' });
		expect(pages('3', 'x')).toEqual({ invalid: 'to' });
	});
});
