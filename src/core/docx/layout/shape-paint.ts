import type { InlineImage } from '../inline-model';
import { drawingShapeSvg, shapeDefaultTextColor } from '../shape-svg';
import type { ThemeCatalog } from '../theme-model';
import type { LayoutObject } from './input';
import type { TextBoxLayouts } from './text-box-layout';

/**
 * A drawing shape as the views paint it: its geometry as SVG (drawn like a chart) and, for a text
 * box, its paragraphs with the colour they take where their runs set none (`shapeDefaultTextColor`,
 * the colour the shape's own text is drawn in). With `textBoxes`, the paragraphs are laid out by
 * the page engine, and a shape that autofit or `wrap="none"` resizes is drawn, and sized, as laid
 * out. Nothing when it cannot be drawn.
 */
export function shapePaint(
	image: Pick<InlineImage, 'shape' | 'textBoxContent' | 'widthPx' | 'heightPx'>,
	theme?: ThemeCatalog,
	textBoxes?: TextBoxLayouts,
): Pick<LayoutObject, 'chart' | 'textBox' | 'textBoxLayout'> &
	Partial<Pick<LayoutObject, 'widthPx' | 'heightPx'>> {
	const layout = image.shape && image.textBoxContent ? textBoxes?.layout(image) : undefined;
	const size = layout ? { widthPx: layout.widthPx, heightPx: layout.heightPx } : {};
	const svg = drawingShapeSvg({ ...image, ...size }, theme);
	if (!svg || !image.shape) return {};
	if (!image.textBoxContent) return { chart: { svg, label: 'Shape', shape: true } };
	const textColor = shapeDefaultTextColor(image.shape, theme);
	return {
		chart: { svg, label: 'Shape', shape: true, textColor },
		textBox: image.textBoxContent,
		...(layout ? { textBoxLayout: layout, ...size } : {}),
	};
}
