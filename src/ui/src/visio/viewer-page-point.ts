import type { VisioPage } from 'ooxml-core/visio';
import { visioDrawingPoint } from 'ooxml-core/visio/ui';

/** Client coordinates projected into physical SVG page inches. */
export function pagePoint(
	svg: SVGSVGElement,
	page: VisioPage,
	event: Pick<MouseEvent, 'clientX' | 'clientY'>,
	options: { snap?: boolean; bounded?: boolean } = {},
) {
	const matrix = svg.getScreenCTM?.();
	if (!matrix) return undefined;
	const point = new DOMPoint(event.clientX, event.clientY).matrixTransform(matrix.inverse());
	return visioDrawingPoint(page, point, options);
}
