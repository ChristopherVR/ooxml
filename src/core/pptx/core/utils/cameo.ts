import type { XmlObject } from '../types';
import { xmlChild } from './xml-access';

const LIVE_FEED_URI = '{51228E76-BA90-4043-B771-695A4F85340A}';
const LIVE_FEED_NAMESPACE = 'http://schemas.microsoft.com/office/drawing/2021/livefeed';

/** [MS-ODRAWXML] §2.42.1.1. Keep the full extension in the picture's raw XML. */
export function isLiveFeedPicture(picture: XmlObject): boolean {
	const nv = xmlChild(picture, 'p:nvPicPr');
	const blip = xmlChild(xmlChild(picture, 'p:blipFill'), 'a:blip');
	for (const container of [
		blip,
		xmlChild(picture, 'p:spPr'),
		xmlChild(nv, 'p:nvPr'),
		xmlChild(nv, 'p:cNvPr'),
		xmlChild(nv, 'p:cNvPicPr'),
		picture,
	]) {
		const list = xmlChild(container, 'a:extLst') ?? xmlChild(container, 'p:extLst');
		const raw = list?.['a:ext'] ?? list?.['p:ext'];
		const extensions = Array.isArray(raw) ? raw : [raw];
		if (
			extensions.some(
				(ext: unknown) =>
					typeof ext === 'object' &&
					ext !== null &&
					'@_uri' in ext &&
					String(ext['@_uri']).toUpperCase() === LIVE_FEED_URI &&
					Object.keys(ext).some((name) => name.endsWith(':liveFeedProps')),
			)
		)
			return true;
	}
	return false;
}

/** An embedded poster remains visible in readers without live camera support. */
export function liveFeedBlipExtension(): XmlObject {
	return {
		'a:ext': {
			'@_uri': LIVE_FEED_URI,
			'alf:liveFeedProps': { '@_xmlns:alf': LIVE_FEED_NAMESPACE },
		},
	};
}
