/**
 * Paints a content control tag's outline as the vector path Word's PDF export writes: both side
 * edges are arcs, filled and stroked with one even-odd `B*`, and paragraph, row and cell tags
 * carry a small arrow and bar. DOM-free.
 */
import {
	contentControlTagMarker,
	contentControlTagOutline,
	contentControlTagOutlineWidth,
	contentControlTagStyle,
	ptToPx,
	type LayoutFragment,
} from 'ooxml-core/docx/layout';
import { parseColor } from '../../export/pdf-svg-style';
import type { PdfVectorPage } from '../../export/pdf-vector';

const BLACK = { r: 0, g: 0, b: 0, a: 1 };
const num = (value: number) => String(Number(value.toFixed(3)));

function fill(page: PdfVectorPage, color: string): void {
	const { r, g, b } = parseColor(color) ?? BLACK;
	page.fillColor(r, g, b);
}

function stroke(page: PdfVectorPage, color: string): void {
	const { r, g, b } = parseColor(color) ?? BLACK;
	page.strokeColor(r, g, b);
}

/** The outline of `fragment`'s tag, whose box starts at (`x`, `top`) in pixels. */
export function paintContentControlTag(
	page: PdfVectorPage,
	fragment: LayoutFragment,
	x: number,
	top: number,
): void {
	const edge = fragment.contentControlTag;
	if (!edge) return;
	const structural = !!fragment.blockContentControlTag;
	const style = contentControlTagStyle;
	const widthPt = contentControlTagOutlineWidth(edge, (fragment.widthPx * 72) / 96, structural);
	const point = (px: number, py: number) => `${num(x + ptToPx(px))} ${num(top + ptToPx(py))}`;
	fill(page, structural ? style.blockFill : style.fill);
	stroke(page, style.stroke);
	page.lineStyle(ptToPx(style.strokeWidthPt), 0, 1);
	for (const command of contentControlTagOutline(edge, widthPt)) {
		if (command[0] === 'C') {
			const [, x1, y1, x2, y2, x3, y3] = command;
			page.op(`${point(x1, y1)} ${point(x2, y2)} ${point(x3, y3)} c`);
		} else page.op(`${point(command[1], command[2])} ${command[0] === 'M' ? 'm' : 'l'}`);
	}
	page.op('h B*');
	if (!structural) return;
	const { arrow, bar } = contentControlTagMarker(edge, widthPt);
	fill(page, style.stroke);
	page.lineStyle(ptToPx(style.marker.strokeWidthPt), 0, 1);
	arrow.forEach(([px, py], index) => page.op(`${point(px, py)} ${index ? 'l' : 'm'}`));
	page.op('h B*');
	page.rect(x + ptToPx(bar.x), top + ptToPx(bar.y), ptToPx(bar.width), ptToPx(bar.height));
	page.op('B*');
}
