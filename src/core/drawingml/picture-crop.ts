import type { CropInsets } from '../geometry/picture-crop';
import { NS, first, type XmlElement } from './dom';

export type DrawingPictureCrop = Partial<CropInsets>;
const SIDES = { l: 'cropLeft', t: 'cropTop', r: 'cropRight', b: 'cropBottom' } as const;

/** Signed DrawingML percentages: integer thousandths of a percent or a strict percentage. */
export function parsePictureCrop(blipFill: XmlElement | undefined): DrawingPictureCrop | undefined {
	const rect = first(blipFill, 'srcRect', NS.a);
	if (!rect) return undefined;
	const result: DrawingPictureCrop = {};
	for (const [attr, key] of Object.entries(SIDES)) {
		const raw = rect.getAttribute(attr);
		if (raw === null || !raw.trim()) continue;
		const percent = raw.endsWith('%');
		const value = Number(percent ? raw.slice(0, -1) : raw) / (percent ? 100 : 100000);
		if (Number.isFinite(value)) result[key] = value;
	}
	return result;
}

/** Patches just srcRect attributes; opaque attributes/children and sibling order survive. */
export function applyPictureCrop(blipFill: XmlElement, crop: DrawingPictureCrop | undefined): void {
	let rect = first(blipFill, 'srcRect', NS.a);
	if (!rect && crop && Object.values(crop).some((value) => value !== undefined)) {
		rect = blipFill.ownerDocument.createElementNS(NS.a, 'a:srcRect');
		const blip = first(blipFill, 'blip', NS.a);
		blipFill.insertBefore(rect, blip?.nextSibling ?? blipFill.firstChild);
	}
	if (!rect) return;
	for (const [attr, key] of Object.entries(SIDES)) {
		const value = crop?.[key];
		if (value === undefined) rect.removeAttribute(attr);
		else {
			if (!Number.isFinite(value)) throw new RangeError('Picture crop must be finite');
			rect.setAttribute(attr, String(Math.round(value * 100000)));
		}
	}
}
