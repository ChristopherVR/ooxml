/** The page's footnote area for the Word PDF export, placed as Print Layout places it. DOM-free. */
import type { LayoutPageBox, LayoutParagraphBox } from 'ooxml-core/docx/layout';
import { paintParagraph, type DocxPdfPainter } from './pdf-blocks';
import { fillWith } from './pdf-text';

/** Word's short separator rule: 1 px, 7 px above the notes, at most 2 inches wide. */
function rule(painter: DocxPdfPainter, x: number, y: number, width: number) {
	if (!fillWith(painter.page, '#000000')) return;
	painter.page.rect(x, y - 7, width, 1);
	painter.page.op('f');
}

function separator(
	painter: DocxPdfPainter,
	story: { heightPx: number; paragraphs: LayoutParagraphBox[] } | undefined,
	x: number,
	y: number,
	width: number,
	continued: boolean,
) {
	if (!story) return rule(painter, x, y, continued ? width : Math.min(192, width));
	for (const paragraph of story.paragraphs)
		paintParagraph(painter, paragraph, x, y - story.heightPx);
}

export function paintFootnotes(painter: DocxPdfPainter, page: LayoutPageBox): void {
	const notes = page.footnotes ?? [];
	if (!notes.length) return;
	const heightPx = page.footnoteHeightPx ?? notes.reduce((sum, note) => sum + note.heightPx, 0);
	const widthPx = page.widthPx - page.marginLeftPx - page.marginRightPx;
	const left = page.marginLeftPx;
	const top = page.heightPx - page.marginBottomPx - heightPx;
	if (page.footnoteColumnAreas)
		for (const area of page.footnoteColumnAreas)
			separator(
				painter,
				area.separator,
				left + area.xPx,
				top + area.yPx,
				area.widthPx,
				area.continued,
			);
	else
		separator(
			painter,
			page.footnoteSeparator,
			left,
			top,
			widthPx,
			notes.some((note) => note.continued),
		);
	for (const note of notes)
		for (const paragraph of note.paragraphs)
			paintParagraph(painter, paragraph, left + (note.xPx ?? 0), top + note.yPx);
}
