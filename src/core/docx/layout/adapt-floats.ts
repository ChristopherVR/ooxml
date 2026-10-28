import type { Paragraph, ThemeCatalog } from '../index';
import { shapeSvg } from '../shape-svg';
import type { LayoutFloat } from './input';

/** Floating pictures anchored in a paragraph, with their `wp:positionH`/`wp:positionV` placement. */
export function paragraphFloats(paragraph: Paragraph, theme?: ThemeCatalog): LayoutFloat[] {
	const floats: LayoutFloat[] = [];
	for (const run of paragraph.runs) {
		const image = run.image;
		if (!image?.anchored || (!image.partName && !image.shape)) continue;
		const placement = image.placement;
		const svg = image.shape && shapeSvg(image.shape, image.widthPx, image.heightPx, theme);
		floats.push({
			index: floats.length,
			...(svg ? { chart: { svg, label: 'Shape', shape: true as const } } : {}),
			partName: image.partName,
			...(image.pictureFill ? { pictureFill: image.pictureFill } : {}),
			contentType: image.contentType,
			widthPx: image.widthPx,
			heightPx: image.heightPx,
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
export function floatsOf(paragraph: Paragraph, theme?: ThemeCatalog): { floats?: LayoutFloat[] } {
	const floats = paragraphFloats(paragraph, theme);
	return floats.length ? { floats } : {};
}
