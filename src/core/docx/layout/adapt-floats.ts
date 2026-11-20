import type { Paragraph, ThemeCatalog } from '../index';
import { docxChartPaint } from '../chart-paint';
import { shapePaint } from './shape-paint';
import type { LayoutFloat } from './input';
import { reviewSourceLength } from './review-display';
import type { TextBoxLayouts } from './text-box-layout';

/**
 * Floating pictures, shapes and charts anchored in a paragraph, with their `wp:positionH`/
 * `wp:positionV` placement; with `textBoxes`, text boxes laid out by the page engine.
 * `pageInCells`: `layoutInCell="0"` puts an object anchored in a table cell on the page (before
 * compatibility mode 15; later modes lay every such object out in its cell).
 */
export function paragraphFloats(
	paragraph: Paragraph,
	theme?: ThemeCatalog,
	textBoxes?: TextBoxLayouts,
	pageInCells = false,
): LayoutFloat[] {
	const floats: LayoutFloat[] = [];
	let sourceOffset = 0;
	for (const run of paragraph.runs) {
		const image = run.image;
		const offset = sourceOffset;
		sourceOffset += reviewSourceLength(run);
		if (!image?.anchored || (!image.partName && !image.shape && !image.chart)) continue;
		const placement = image.placement;
		floats.push({
			index: floats.length,
			sourceOffset: offset,
			...(image.chart
				? {
						chart: docxChartPaint(image.chart, image.widthPx, image.heightPx, theme, image.altText),
					}
				: {}),
			partName: image.partName,
			...(image.pictureFill ? { pictureFill: image.pictureFill } : {}),
			contentType: image.contentType,
			widthPx: image.widthPx,
			heightPx: image.heightPx,
			...shapePaint(image, theme, textBoxes),
			...(placement?.relativeFrom ? { relativeFromH: placement.relativeFrom } : {}),
			...(placement?.align ? { alignH: placement.align } : {}),
			...(placement?.offsetXPx !== undefined ? { offsetXPx: placement.offsetXPx } : {}),
			...(placement?.relativeFromV ? { relativeFromV: placement.relativeFromV } : {}),
			...(placement?.alignV ? { alignV: placement.alignV } : {}),
			...(placement?.offsetYPx !== undefined ? { offsetYPx: placement.offsetYPx } : {}),
			...(placement?.behindText ? { behindText: true } : {}),
			...(placement?.wrap ? { wrap: placement.wrap } : {}),
			...(placement?.wrapText ? { wrapText: placement.wrapText } : {}),
			...(pageInCells && placement?.layoutInCell === false ? { layoutInCell: false } : {}),
		});
	}
	return floats;
}
export function floatsOf(
	paragraph: Paragraph,
	theme?: ThemeCatalog,
	textBoxes?: TextBoxLayouts,
	pageInCells = false,
): { floats?: LayoutFloat[] } {
	const floats = paragraphFloats(paragraph, theme, textBoxes, pageInCells);
	return floats.length ? { floats } : {};
}
