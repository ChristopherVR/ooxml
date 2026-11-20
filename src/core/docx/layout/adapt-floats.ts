import type { Paragraph, ThemeCatalog } from '../index';
import { shapePaint } from './shape-paint';
import type { LayoutFloat } from './input';
import type { TextBoxLayouts } from './text-box-layout';

/**
 * Floating pictures anchored in a paragraph, with their `wp:positionH`/`wp:positionV` placement;
 * with `textBoxes`, text boxes laid out by the page engine.
 */
export function paragraphFloats(
	paragraph: Paragraph,
	theme?: ThemeCatalog,
	textBoxes?: TextBoxLayouts,
): LayoutFloat[] {
	const floats: LayoutFloat[] = [];
	for (const run of paragraph.runs) {
		const image = run.image;
		if (!image?.anchored || (!image.partName && !image.shape)) continue;
		const placement = image.placement;
		floats.push({
			index: floats.length,
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
		});
	}
	return floats;
}
export function floatsOf(
	paragraph: Paragraph,
	theme?: ThemeCatalog,
	textBoxes?: TextBoxLayouts,
): { floats?: LayoutFloat[] } {
	const floats = paragraphFloats(paragraph, theme, textBoxes);
	return floats.length ? { floats } : {};
}
