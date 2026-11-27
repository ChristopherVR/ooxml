/**
 * The title block every drawing keeps on a background page: organisation, document title, the
 * name and number of the sheet in front of it (text fields), and a footer line.
 */
import type { Drawing } from './drawing';
import type { PageBuilder } from './page';
import { RUN } from './palette';
import type { Rich } from './text';

export const PAGE = { w: 17, h: 11, left: 0.6, right: 16.4, top: 9.55, bottom: 0.85 } as const;
export const TITLE_BLOCK = 'Title block';

export interface TitleBlock {
	organisation: string;
	title: string;
	/** What the document is, shown in the footer. */
	document: string;
	revision: string;
	/** Style of the small square beside the organisation name. */
	accent: string;
	/** How many foreground pages the drawing has. */
	sheets: number;
}

/** A labelled value in the title block: a small caps label over the value. */
function entry(page: PageBuilder, left: number, w: number, label: string, value: Rich) {
	page.note('Overline', label, { left, y: 10.42, w, h: 0.2 });
	return page.note('Block value', value, { left, y: 10.14, w, h: 0.26 });
}

export async function drawTitleBlock(drawing: Drawing, block: TitleBlock): Promise<void> {
	const page = drawing.page(TITLE_BLOCK);
	page.block(block.accent, { left: PAGE.left, bottom: 10.33, w: 0.18, h: 0.18 });
	page.note('Overline', block.organisation, { left: PAGE.left + 0.28, y: 10.42, w: 6, h: 0.2 });
	page.note('Title', block.title, { left: PAGE.left, y: 10.08, w: 9, h: 0.44, name: 'Title' });
	const sheet = entry(page, 11.4, 2.5, 'Sheet', '');
	// On a background page these functions show the page in front of it. The library cannot
	// evaluate the two page counters there (a background page has no number), so they are
	// written here with the first page as the cached text.
	entry(page, 14.0, 1.0, 'Page', [
		{ text: '1', field: 'PAGENUMBER()', picture: '0', style: RUN.value },
		{ text: ' of ', style: RUN.value },
		{ text: String(block.sheets), field: 'PAGECOUNT()', picture: '0', style: RUN.value },
	]);
	entry(page, 15.1, 1.3, 'Revision', block.revision);
	page.block('Hairline', { left: PAGE.left, bottom: 9.78, w: PAGE.right - PAGE.left, h: 0.012 });
	page.block('Hairline', { left: PAGE.left, bottom: 0.6, w: PAGE.right - PAGE.left, h: 0.012 });
	page.note('Caption', block.document, { left: PAGE.left, y: 0.42, w: 9, h: 0.22 });
	page.note('Caption right', 'Fictitious sample content', {
		left: PAGE.right - 5,
		y: 0.42,
		w: 5,
		h: 0.22,
	});
	await drawing.place(page);
	await drawing.edit(
		[
			{
				type: 'insert-text-field',
				pageId: page.pageId,
				shapeId: String(sheet.id),
				formula: 'PAGENAME()',
				format: '@',
			},
		],
		'sheet name field',
	);
}
