import type { InlineImage } from '../inline-model';
import { drawingShapeSvg, shapeDefaultTextColor } from '../shape-svg';
import type { ThemeCatalog } from '../theme-model';
import type { LayoutObject } from './input';

/**
 * A drawing shape as the views paint it: its geometry as SVG (drawn like a chart) and, for a text
 * box, its paragraphs with the colour they take where their runs set none (`shapeDefaultTextColor`,
 * the colour the shape's own text is drawn in). Nothing when it cannot be drawn.
 */
export function shapePaint(
	image: Pick<InlineImage, 'shape' | 'textBoxContent' | 'widthPx' | 'heightPx'>,
	theme?: ThemeCatalog,
): Pick<LayoutObject, 'chart' | 'textBox'> {
	const svg = drawingShapeSvg(image, theme);
	if (!svg || !image.shape) return {};
	if (!image.textBoxContent) return { chart: { svg, label: 'Shape', shape: true } };
	const textColor = shapeDefaultTextColor(image.shape, theme);
	return {
		chart: { svg, label: 'Shape', shape: true, textColor },
		textBox: image.textBoxContent,
	};
}
