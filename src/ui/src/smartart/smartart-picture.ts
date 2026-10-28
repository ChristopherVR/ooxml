import type { DrawingFill } from 'ooxml-core/drawingml';
import { picturePaint, type PictureMedia, type PicturePaintContext } from '../picture-paint';
import type { SmartArtRenderReport } from './smartart-svg';

export type DiagramPictureResolver = (partName: string) => string | PictureMedia | undefined;
/** Shared picture paint server for source crop, ordered effects and tile layout. */
export function diagramPictureFill(
	doc: Document,
	defs: SVGDefsElement,
	fill: Extract<DrawingFill, { kind: 'picture' }>,
	part: string | undefined,
	resolve: DiagramPictureResolver | undefined,
	report: SmartArtRenderReport,
	context?: PicturePaintContext,
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
	const effects = picturePaint(doc, defs, fill, url, context);
	if (effects.unapplied.length)
		report.unappliedEffects = [
			...new Set([
				...(report.unappliedEffects ?? []),
				...effects.unapplied.map((name) => `picture:${name}`),
			]),
		];
	if (effects.approximated?.length)
		report.approximatedFills.push(...effects.approximated.map((name) => `picture:${name}`));
	return effects.paint;
}
