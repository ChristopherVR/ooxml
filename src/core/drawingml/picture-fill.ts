import { parsePictureLayout } from './picture-layout';
import { NS, buildXml, first, type XmlElement } from '../xml/index';
import type { DrawingFill } from './types';
import { parsePictureCrop } from './picture-crop';
import { parsePictureEffects } from './picture-effects';
/** Pictures and picture fills share blip, source crop and image effects; outer namespace is host-owned. */
export function parseDrawingPictureFill(
	fill: XmlElement,
): Extract<DrawingFill, { kind: 'picture' }> {
	const blip = first(fill, 'blip', NS.a);
	const relId = blip?.getAttributeNS(NS.r, 'embed') || blip?.getAttribute('r:embed') || undefined;
	const linkRelId = blip?.getAttributeNS(NS.r, 'link') || undefined;
	const crop = parsePictureCrop(fill);
	const effects = parsePictureEffects(blip);
	const layout = parsePictureLayout(fill);
	const unsupported: string[] = [];
	const tile = first(fill, 'tile', NS.a),
		stretch = first(fill, 'stretch', NS.a);
	if (tile && stretch) unsupported.push('tile/stretch-choice');
	if (tile) {
		for (const [attr, key] of [
			['sx', 'scaleX'],
			['sy', 'scaleY'],
			['tx', 'offsetXEmu'],
			['ty', 'offsetYEmu'],
		] as const)
			if (
				tile.hasAttribute(attr) &&
				(layout.tile?.[key] === undefined ||
					((attr === 'sx' || attr === 'sy') && layout.tile[key]! <= 0))
			)
				unsupported.push(`tile:${attr}`);
		if (
			tile.hasAttribute('algn') &&
			!['tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br'].includes(tile.getAttribute('algn') ?? '')
		)
			unsupported.push('tile:algn');
		if (tile.hasAttribute('flip') && !layout.tile?.flip) unsupported.push('tile:flip');
	}
	const rect = first(stretch, 'fillRect', NS.a);
	if (rect)
		for (const key of ['l', 't', 'r', 'b']) {
			const raw = rect.getAttribute(key);
			if (raw !== null && !/^[-+]?(?:\d+\.?\d*|\.\d+)%?$/.test(raw))
				unsupported.push(`fillRect:${key}`);
		}
	if (
		fill.hasAttribute('rotWithShape') &&
		!['0', '1', 'true', 'false'].includes(fill.getAttribute('rotWithShape') ?? '')
	)
		unsupported.push('rotWithShape');
	return {
		kind: 'picture',
		sourceXml: buildXml(fill),
		...(relId ? { relId } : {}),
		...(linkRelId ? { linkRelId } : {}),
		...(crop ? { crop } : {}),
		...(effects.length ? { effects } : {}),
		...(Object.keys(layout).length ? { layout } : {}),
		...(unsupported.length ? { unsupportedPicture: unsupported } : {}),
	};
}
