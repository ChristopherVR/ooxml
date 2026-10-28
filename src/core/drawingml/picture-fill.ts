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
	const tile = first(fill, 'tile', NS.a);
	const rect = first(first(fill, 'stretch', NS.a), 'fillRect', NS.a);
	const unsupported = [
		...(tile ? ['tile'] : []),
		...(rect && ['l', 't', 'r', 'b'].some((key) => Number(rect.getAttribute(key) ?? 0) !== 0)
			? ['fillRect']
			: []),
		...(fill.getAttribute('rotWithShape') === '0' || fill.getAttribute('rotWithShape') === 'false'
			? ['rotWithShape']
			: []),
	];
	return {
		kind: 'picture',
		sourceXml: buildXml(fill),
		...(relId ? { relId } : {}),
		...(linkRelId ? { linkRelId } : {}),
		...(crop ? { crop } : {}),
		...(effects.length ? { effects } : {}),
		...(unsupported.length ? { unsupportedPicture: unsupported } : {}),
	};
}
