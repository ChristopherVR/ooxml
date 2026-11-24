/**
 * Paints laid-out paragraphs and tables as PDF operators, with the geometry Print Layout gives
 * them (`print-layout.ts`, `print-table.ts`): paragraph borders and shading, cell shading and
 * borders inside each cell's box, cell content inset by its border and padding and aligned
 * vertically, clipped to the cell. Coordinates are CSS pixels, y down. DOM-free.
 */
import {
	cellVerticalOffset,
	type LayoutBlockBox,
	type LayoutBorder,
	type LayoutParagraphBox,
	type LayoutParagraphFrame,
	type LayoutTableBox,
} from 'ooxml-core/docx/layout';
import type { PdfVectorPage } from '../../export/pdf-vector';
import type { DocxPdfPictures } from './pdf-pictures';
import { cssRgba, fillWith, type DocxPdfText } from './pdf-text';

export interface DocxPdfPainter {
	page: PdfVectorPage;
	text: DocxPdfText;
	pictures: DocxPdfPictures;
}

const borderWidth = (border: LayoutBorder | undefined) => (border ? border.widthPx : 0);

/** One border side as filled rectangles (double lines as two, dashes and dots as a stroke). */
export function drawBorder(
	page: PdfVectorPage,
	border: LayoutBorder | undefined,
	x: number,
	y: number,
	length: number,
	horizontal: boolean,
): void {
	if (!border || !(border.widthPx > 0) || !(length > 0)) return;
	const rgba = cssRgba(border.color) ?? cssRgba('#000000')!;
	if (rgba.a === 0) return;
	const width = border.widthPx;
	if (border.style === 'dashed' || border.style === 'dotted') {
		const dash = border.style === 'dotted' ? [width, width] : [width * 3, width * 2];
		page.strokeColor(rgba.r, rgba.g, rgba.b);
		page.lineStyle(width, 0, 0, dash);
		const middle = width / 2;
		const n = (value: number) => Number(value.toFixed(3));
		page.op(
			horizontal
				? `${n(x)} ${n(y + middle)} m ${n(x + length)} ${n(y + middle)} l S`
				: `${n(x + middle)} ${n(y)} m ${n(x + middle)} ${n(y + length)} l S`,
		);
		return;
	}
	page.fillColor(rgba.r, rgba.g, rgba.b);
	const bands =
		border.style === 'double' && width >= 3
			? [
					[0, width / 3],
					[(2 * width) / 3, width / 3],
				]
			: [[0, width]];
	for (const [offset, size] of bands)
		if (horizontal) page.rect(x, y + offset!, length, size!);
		else page.rect(x + offset!, y, size!, length);
	page.op('f');
}

function paragraphFrame(
	page: PdfVectorPage,
	frame: LayoutParagraphFrame,
	x: number,
	y: number,
	height: number,
) {
	const left = x + frame.leftPx;
	if (frame.shading && fillWith(page, frame.shading, 'transparent')) {
		page.rect(left, y, frame.widthPx, height);
		page.op('f');
	}
	const borders = frame.borders;
	if (!borders) return;
	// Borders sit inside the frame's box, as the CSS border-box frame draws them.
	drawBorder(page, borders.top, left, y, frame.widthPx, true);
	drawBorder(
		page,
		borders.bottom,
		left,
		y + height - borderWidth(borders.bottom),
		frame.widthPx,
		true,
	);
	drawBorder(page, borders.left, left, y, height, false);
	drawBorder(
		page,
		borders.right,
		left + frame.widthPx - borderWidth(borders.right),
		y,
		height,
		false,
	);
}

export function paintParagraph(
	painter: DocxPdfPainter,
	box: LayoutParagraphBox,
	originX: number,
	originY: number,
): void {
	const { page, text, pictures } = painter;
	const x = originX + (box.xPx ?? 0);
	const y = originY + box.yPx;
	if (box.frame) paragraphFrame(page, box.frame, x, y, box.heightPx);
	if (box.floats?.length)
		painter.text.report.omitted(
			'picture',
			'Floating pictures inside table cells are not exported to PDF yet.',
		);
	for (const line of box.lines) {
		const lineTop = y + line.yPx;
		line.fragments.forEach((fragment, index) => {
			if (fragment.object) {
				pictures.draw(page, fragment.object, x + fragment.xPx, lineTop + (fragment.topPx ?? 0));
				return;
			}
			if (!fragment.text && !fragment.leader && !fragment.background) return;
			// A space fragment after a word is written with the word as its trailing space, stretched to
			// the space's laid-out width (justified lines widen it); its own fragment then draws only
			// its highlight and lines.
			const next = line.fragments[index + 1];
			const spaced = !!fragment.text.trim() && !!next?.text && !next.object && !next.text.trim();
			const gap = spaced
				? Math.max(0, next.xPx + next.widthPx - fragment.xPx - fragment.widthPx)
				: 0;
			text.draw(page, fragment, x + fragment.xPx, line, lineTop, gap);
		});
	}
}

export function paintTable(
	painter: DocxPdfPainter,
	box: LayoutTableBox,
	originX: number,
	originY: number,
	withText = true,
): void {
	const { page } = painter;
	const tableX = originX + (box.xPx ?? 0);
	const tableY = originY + box.yPx;
	box.rows.forEach((row, rowIndex) => {
		const rowY = tableY + row.yPx;
		const lastRow = rowIndex === box.rows.length - 1;
		row.cells.forEach((paragraphs, cellIndex) => {
			const geometry = row.geometry?.[cellIndex] ?? {
				xPx: (cellIndex * (box.widthPx ?? 0)) / Math.max(1, row.cells.length),
				widthPx: (box.widthPx ?? 0) / Math.max(1, row.cells.length),
				paddingLeftPx: 0,
				paddingRightPx: 0,
				contentHeightPx: 0,
			};
			const x = tableX + geometry.xPx;
			const width = geometry.widthPx;
			const height = row.heightPx;
			if (geometry.shading && fillWith(page, geometry.shading, 'transparent')) {
				page.rect(x, rowY, width, height);
				page.op('f');
			}
			const borders = geometry.borders;
			const lastCell = cellIndex === row.cells.length - 1;
			drawBorder(page, borders?.top, x, rowY, width, true);
			drawBorder(page, borders?.left, x, rowY, height, false);
			if (lastCell)
				drawBorder(
					page,
					borders?.right,
					x + width - borderWidth(borders?.right),
					rowY,
					height,
					false,
				);
			if (lastRow)
				drawBorder(
					page,
					borders?.bottom,
					x,
					rowY + height - borderWidth(borders?.bottom),
					width,
					true,
				);
			const offset = cellVerticalOffset(geometry, row.heightPx);
			const reserved = geometry.borderInsetsReserved;
			const contentX = x + (reserved ? 0 : borderWidth(borders?.left)) + geometry.paddingLeftPx;
			const contentY = rowY + (reserved ? 0 : borderWidth(borders?.top)) + offset;
			for (const nested of row.cellTables?.[cellIndex] ?? [])
				paintTable(painter, nested, x + geometry.paddingLeftPx, rowY + offset, false);
			if (!withText || !paragraphs.length) return;
			// A vertical merge shows its text below its first row; other cells clip to their box.
			const clip = !(geometry.mergeRows && geometry.mergeRows > 1);
			if (clip) {
				page.save();
				page.rect(x, rowY, width, height);
				page.op('W n');
			}
			for (const paragraph of paragraphs) paintParagraph(painter, paragraph, contentX, contentY);
			if (clip) page.restore();
		});
	});
}

export function paintBlock(
	painter: DocxPdfPainter,
	box: LayoutBlockBox,
	originX: number,
	originY: number,
): void {
	if (box.kind === 'paragraph') paintParagraph(painter, box, originX, originY);
	else paintTable(painter, box, originX, originY);
}
