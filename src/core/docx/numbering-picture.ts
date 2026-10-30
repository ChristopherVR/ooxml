// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import { imageContentType } from '../drawingml/image-content-type';
import { resolvePartPath, type Relationship } from '../opc/relationships';
import type { NumberingPictureBullet } from './numbering-model';
import { getR, getW, type XmlElement } from './xml';

/** The numbering part's relationships, used to resolve picture bullet images. */
export type NumberingRelationships = ReadonlyMap<string, Relationship>;
export const NUMBERING_PART = 'word/numbering.xml';

const descendants = (root: XmlElement, local: string): XmlElement[] =>
	Array.from(root.getElementsByTagName('*')).filter(
		(node): node is XmlElement => node.nodeType === 1 && (node as XmlElement).localName === local,
	);

const points = (style: string, property: string): number | undefined => {
	const match = new RegExp(`(?:^|;)\\s*${property}:\\s*([\\d.]+)(pt|px|in|cm|mm)?`, 'i').exec(
		style,
	);
	if (!match) return undefined;
	const value = Number(match[1]);
	const unit = (match[2] ?? 'pt').toLowerCase();
	const factor = { pt: 1, px: 0.75, in: 72, cm: 72 / 2.54, mm: 72 / 25.4 }[unit] ?? 1;
	return Number.isFinite(value) && value > 0 ? value * factor : undefined;
};

function parsePictureBullet(element: XmlElement, rels: NumberingRelationships) {
	const id = getW(element, 'numPicBulletId');
	if (id === undefined) return undefined;
	// Older files: w:pict/v:shape/v:imagedata[r:id]; newer ones: w:drawing with an a:blip[r:embed].
	const imagedata = descendants(element, 'imagedata')[0];
	const blip = descendants(element, 'blip')[0];
	const relId = (imagedata && getR(imagedata, 'id')) || (blip && getR(blip, 'embed')) || undefined;
	const rel = relId ? rels.get(relId) : undefined;
	if (!rel || rel.mode === 'External') return undefined;
	const partName = resolvePartPath(NUMBERING_PART, rel.target);
	const result: NumberingPictureBullet = {
		id,
		partName,
		contentType: imageContentType(partName),
	};
	const style = descendants(element, 'shape')[0]?.getAttribute('style') ?? '';
	let width = points(style, 'width');
	let height = points(style, 'height');
	const extent = descendants(element, 'extent')[0];
	if (extent && width === undefined && height === undefined) {
		const cx = Number(extent.getAttribute('cx'));
		const cy = Number(extent.getAttribute('cy'));
		if (cx > 0 && cy > 0) {
			width = cx / 12700;
			height = cy / 12700;
		}
	}
	if (width !== undefined) result.widthPt = width;
	if (height !== undefined) result.heightPt = height;
	return result;
}

/** Reads every `w:numPicBullet` whose image resolves through the numbering relationships. */
export function parseNumberingPictureBullets(
	root: XmlElement,
	rels: NumberingRelationships | undefined,
): Record<string, NumberingPictureBullet> {
	const bullets: Record<string, NumberingPictureBullet> = Object.create(null);
	if (!rels) return bullets;
	for (const element of descendants(root, 'numPicBullet')) {
		const parsed = parsePictureBullet(element, rels);
		if (parsed) bullets[parsed.id] = parsed;
	}
	return bullets;
}

/**
 * The drawn size of a picture bullet in points. Word sizes the picture to the marker's font size
 * unless the shape authors its own width and height, which are honoured (a missing side keeps the
 * picture square).
 */
export function pictureBulletSizePt(
	picture: Pick<NumberingPictureBullet, 'widthPt' | 'heightPt'>,
	fontSizePt: number,
): { widthPt: number; heightPt: number } {
	const heightPt = picture.heightPt ?? picture.widthPt ?? fontSizePt;
	return { widthPt: picture.widthPt ?? heightPt, heightPt };
}
