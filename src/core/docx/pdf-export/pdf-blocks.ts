/**
 * Paints laid-out paragraphs and tables as PDF operators, with the geometry Print Layout gives
 * them (`print-layout.ts`, `print-table.ts`): paragraph borders and shading, cell shading and
 * borders inside each cell's box, cell content inset by its border and padding and aligned
 * vertically, clipped to the cell. Coordinates are CSS pixels, y down. DOM-free.
 */
import {
	cellVerticalOffset,
	paragraphBorderBands,
	paragraphFrameBounds,
	type LayoutBlockBox,
	type LayoutBorder,
	type LayoutParagraphBox,
	type LayoutParagraphFrame,
	type LayoutTableBox,
} from '../layout/index';
import type { PdfVectorPage } from '../../pdf/write/pdf-vector';
import type { DocxPdfHooks } from './pdf-hooks';
import type { DocxPdfPictures } from './pdf-pictures';
import { cssRgba, fillWith, type DocxPdfText } from './pdf-text';

export interface DocxPdfPainter {
	page: PdfVectorPage;
	text: DocxPdfText;
	pictures: DocxPdfPictures;
	/** Tags, links and markup (`pdf-hooks.ts`); absent, blocks are only painted. */
	hooks?: DocxPdfHooks;
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

/** Shading and borders inside a box `frame.widthPx` wide and `height` tall (a CSS border box). */
export function paragraphFrame(
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
	for (const side of ['top', 'bottom', 'left', 'right'] as const) {
		const line = borders[side];
		if (!line) continue;
		const horizontal = side === 'top' || side === 'bottom';
		const reverse = side === 'right' || side === 'bottom';
		const length = horizontal ? frame.widthPx : height;
		const edge = horizontal ? y : left;
		const start = edge + (reverse ? (horizontal ? height : frame.widthPx) - line.widthPx : 0);
		if (!line.compound) {
			drawBorder(page, line, horizontal ? left : start, horizontal ? start : y, length, horizontal);
			continue;
		}
		for (const { offsetPx, widthPx } of paragraphBorderBands(line, reverse)) {
			const offset = start + offsetPx;
			drawBorder(
				page,
				{ ...line, widthPx },
				horizontal ? left : offset,
				horizontal ? offset : y,
				length,
				horizontal,
			);
		}
	}
}

export function paintParagraph(
	painter: DocxPdfPainter,
	box: LayoutParagraphBox,
	originX: number,
	originY: number,
): void {
	const { page, text, pictures, hooks } = painter;
	const x = originX + (box.xPx ?? 0);
	const y = originY + box.yPx;
	// Objects floating in table cells (`box.floats`) are on the page's floats, drawn with them.
	if (box.frame) {
		hooks?.decoration();
		const { topPx, heightPx } = paragraphFrameBounds(box.frame, box.heightPx);
		paragraphFrame(page, box.frame, x, y + topPx, heightPx);
	}
	hooks?.beginParagraph(box, x, y);
	for (const line of box.lines) {
		const lineTop = y + line.yPx;
		line.fragments.forEach((laid, index) => {
			if (laid.contentControlTag) {
				hooks?.decoration();
				text.draw(page, laid, x + laid.xPx, line, lineTop);
				return;
			}
			const fragment = hooks ? hooks.fragment(box, laid, x + laid.xPx, lineTop, line) : laid;
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
	hooks?.endParagraph();
}

type TableRow = LayoutTableBox['rows'][number];

function cellGeometry(box: LayoutTableBox, row: TableRow, cellIndex: number) {
	return (
		row.geometry?.[cellIndex] ?? {
			xPx: (cellIndex * (box.widthPx ?? 0)) / Math.max(1, row.cells.length),
			widthPx: (box.widthPx ?? 0) / Math.max(1, row.cells.length),
			paddingLeftPx: 0,
			paddingRightPx: 0,
			contentHeightPx: 0,
		}
	);
}

/**
 * Every cell's shading, before any row's text: a vertically merged cell draws its text down over
 * the rows it spans, and a later row's shading painted after it would cover that text.
 */
function paintCellShading(
	painter: DocxPdfPainter,
	box: LayoutTableBox,
	tableX: number,
	tableY: number,
): void {
	const { page, hooks } = painter;
	let started = false;
	for (const row of box.rows) {
		row.cells.forEach((_cell, cellIndex) => {
			const geometry = cellGeometry(box, row, cellIndex);
			if (!geometry.shading || !fillWith(page, geometry.shading, 'transparent')) return;
			if (!started) hooks?.decoration();
			started = true;
			page.rect(tableX + geometry.xPx, tableY + row.yPx, geometry.widthPx, row.heightPx);
			page.op('f');
		});
	}
	if (started) hooks?.close();
}

export function paintTable(
	painter: DocxPdfPainter,
	box: LayoutTableBox,
	originX: number,
	originY: number,
	withText = true,
): void {
	const { page, hooks } = painter;
	const tableX = originX + (box.xPx ?? 0);
	const tableY = originY + box.yPx;
	hooks?.beginTable(box.blockId, withText, tableX, tableY, box.heightPx);
	paintCellShading(painter, box, tableX, tableY);
	box.rows.forEach((row, rowIndex) => {
		const rowY = tableY + row.yPx;
		const lastRow = rowIndex === box.rows.length - 1;
		hooks?.beginRow(row, rowIndex, rowY);
		row.cells.forEach((paragraphs, cellIndex) => {
			hooks?.beginCell(cellIndex, tableX, rowY, row.heightPx);
			hooks?.decoration();
			const geometry = cellGeometry(box, row, cellIndex);
			const x = tableX + geometry.xPx;
			const width = geometry.widthPx;
			const height = row.heightPx;
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
			if (withText && paragraphs.length) {
				// A vertical merge shows its text below its first row; other cells clip to their box.
				const clip = !(geometry.mergeRows && geometry.mergeRows > 1);
				hooks?.close();
				if (clip) {
					page.save();
					page.rect(x, rowY, width, height);
					page.op('W n');
				}
				for (const paragraph of paragraphs) paintParagraph(painter, paragraph, contentX, contentY);
				hooks?.close();
				if (clip) page.restore();
			}
			hooks?.endCell();
		});
		hooks?.endRow();
	});
	hooks?.endTable();
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
