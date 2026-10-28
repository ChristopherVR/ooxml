import type { DrawingFill } from 'ooxml-core/drawingml';
import { picturePaint, type PicturePaintContext } from './picture-paint';
export function pictureNeedsSvg(
	fill: Extract<DrawingFill, { kind: 'picture' }> | undefined,
): boolean {
	return Boolean(
		fill &&
		(fill.crop ||
			fill.layout?.tile ||
			Object.values(fill.layout?.fillRect ?? {}).some((n) => n !== 0) ||
			fill.effects?.some((effect) => effect.kind !== 'unsupported')),
	);
}
export function pictureSvg(
	doc: Document,
	fill: Extract<DrawingFill, { kind: 'picture' }>,
	url: string,
	width: number,
	height: number,
	context: Partial<PicturePaintContext> = {},
): SVGSVGElement {
	const ns = 'http://www.w3.org/2000/svg';
	const svg = doc.createElementNS(ns, 'svg');
	svg.setAttribute('width', String(width));
	svg.setAttribute('height', String(height));
	svg.setAttribute('viewBox', `0 0 ${width} ${height}`);
	svg.style.width = '100%';
	svg.style.height = '100%';
	const defs = doc.createElementNS(ns, 'defs');
	svg.append(defs);
	const result = picturePaint(doc, defs, fill, url, { width, height, ...context });
	const rect = doc.createElementNS(ns, 'rect');
	rect.setAttribute('width', String(width));
	rect.setAttribute('height', String(height));
	rect.setAttribute('fill', result.paint);
	svg.append(rect);
	if (result.unapplied.length) svg.dataset.unappliedPictureEffects = result.unapplied.join(',');
	if (result.approximated?.length)
		svg.dataset.approximatedPictureEffects = result.approximated.join(',');
	return svg;
}
