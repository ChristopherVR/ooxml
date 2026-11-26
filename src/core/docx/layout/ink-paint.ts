import { inkSvg } from '../../drawingml/inkml/svg';
import type { InlineImage } from '../inline-model';
import type { LayoutObject } from './input';

/**
 * Ink as the views paint it: its strokes as SVG, drawn like a shape (past its box, no backdrop),
 * stretched to the drawing's size. Nothing for a drawing that is not ink.
 */
export function inkPaint(
	image: Pick<InlineImage, 'ink' | 'widthPx' | 'heightPx'>,
): Pick<LayoutObject, 'chart'> {
	const ink = image.ink;
	if (!ink) return {};
	return {
		chart: {
			svg: inkSvg(ink.strokes, image.widthPx, image.heightPx, ink.width, ink.height),
			label: 'Ink',
			shape: true,
		},
	};
}
