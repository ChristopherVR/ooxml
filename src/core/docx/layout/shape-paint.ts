import type { InlineImage } from '../inline-model';
import { drawingShapeSvg, shapeAppearance, shapeTextColor } from '../shape-svg';
import type { ThemeCatalog } from '../theme-model';
import type { LayoutObject } from './input';

/**
 * A drawing shape as the views paint it: its geometry as SVG (drawn like a chart) and, for a text
 * box, its paragraphs with the colour they take where their runs set none (white on the default and
 * dark fills, black on light ones, as the shape's own text is drawn). Nothing when it cannot be drawn.
 */
export function shapePaint(
	image: Pick<InlineImage, 'shape' | 'textBoxContent' | 'widthPx' | 'heightPx'>,
	theme?: ThemeCatalog,
): Pick<LayoutObject, 'chart' | 'textBox'> {
	const svg = drawingShapeSvg(image, theme);
	if (!svg || !image.shape) return {};
	if (!image.textBoxContent) return { chart: { svg, label: 'Shape', shape: true } };
	const textColor = shapeTextColor(shapeAppearance(image.shape, theme).fill);
	return {
		chart: { svg, label: 'Shape', shape: true, textColor },
		textBox: image.textBoxContent,
	};
}
