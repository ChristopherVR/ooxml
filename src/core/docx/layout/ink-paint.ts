import { inkSvg } from '../../drawingml/inkml/svg';
import type { InlineImage } from '../inline-model';
import type { LayoutObject } from './input';

/** The painted SVG of each ink drawing, by size: a layout paints every drawing again. */
const painted = new WeakMap<object, Map<string, string>>();

/**
 * Ink as the views paint it: its strokes as SVG, drawn like a shape (past its box, no backdrop),
 * stretched to the drawing's size, and marked `ink` for the Draw tab's tools. Nothing for a
 * drawing that is not ink.
 */
export function inkPaint(
	image: Pick<InlineImage, 'ink' | 'widthPx' | 'heightPx'>,
): Pick<LayoutObject, 'chart'> {
	const ink = image.ink;
	if (!ink) return {};
	const sizes = painted.get(ink) ?? new Map<string, string>();
	painted.set(ink, sizes);
	const key = `${image.widthPx}x${image.heightPx}`;
	let svg = sizes.get(key);
	if (svg === undefined) {
		svg = inkSvg(ink.strokes, image.widthPx, image.heightPx, ink.width, ink.height);
		if (sizes.size >= 4) sizes.clear();
		sizes.set(key, svg);
	}
	return { chart: { svg, label: 'Ink', shape: true, ink: true } };
}
