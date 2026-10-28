import { imageRectFromCrop } from 'ooxml-core/geometry';
import type { DrawingFill } from 'ooxml-core/drawingml';
import { pictureEffectFilter } from '../picture-effects';
import type { SmartArtRenderReport } from './smartart-svg';

export type DiagramPictureResolver = (partName: string) => string | undefined;
const NS = 'http://www.w3.org/2000/svg';
let serial = 0;
/** A clipped, source-cropped picture paint server; tiled/offset fills stay explicit fallbacks. */
export function diagramPictureFill(
	doc: Document,
	defs: SVGDefsElement,
	fill: Extract<DrawingFill, { kind: 'picture' }>,
	part: string | undefined,
	resolve: DiagramPictureResolver | undefined,
	report: SmartArtRenderReport,
): string | undefined {
	const fail = (reason: string) => {
		const key = `picture:${reason}`;
		if (!report.approximatedFills.includes(key)) report.approximatedFills.push(key);
		return undefined;
	};
	if (fill.unsupportedPicture?.length) return fail(fill.unsupportedPicture.join(','));
	if (fill.linkRelId && !fill.relId) return fail('external');
	if (!part || !resolve) return fail('unresolved');
	const url = resolve(part);
	if (!url) return fail('missing-media');
	const crop = { cropLeft: 0, cropTop: 0, cropRight: 0, cropBottom: 0, ...fill.crop };
	if (1 - crop.cropLeft - crop.cropRight <= 0 || 1 - crop.cropTop - crop.cropBottom <= 0)
		return fail('invalid-crop');
	const rect = imageRectFromCrop({ x: 0, y: 0, width: 1, height: 1 }, crop);
	const pattern = doc.createElementNS(NS, 'pattern');
	pattern.id = `office-smartart-picture-${++serial}`;
	pattern.setAttribute('width', '1');
	pattern.setAttribute('height', '1');
	pattern.setAttribute('patternContentUnits', 'objectBoundingBox');
	const image = doc.createElementNS(NS, 'image');
	for (const key of ['x', 'y', 'width', 'height'] as const)
		image.setAttribute(key, String(rect[key]));
	image.setAttribute('href', url);
	image.setAttribute('preserveAspectRatio', 'none');
	const effects = pictureEffectFilter(doc, defs, fill.effects ?? []);
	if (effects.filter) image.setAttribute('filter', effects.filter);
	if (effects.unapplied.length)
		report.unappliedEffects = [
			...new Set([
				...(report.unappliedEffects ?? []),
				...effects.unapplied.map((name) => `picture:${name}`),
			]),
		];
	pattern.append(image);
	defs.append(pattern);
	return `url(#${pattern.id})`;
}
