import type { DrawingTransform } from './drawing-arrange';
import type { XmlElement } from './xml';

export function readDrawingTransform(xfrm: XmlElement | undefined): DrawingTransform | undefined {
	if (!xfrm) return undefined;
	const rotation = Number(xfrm.getAttribute('rot')) / 60000;
	const flipH = /^(1|true)$/.test(xfrm.getAttribute('flipH') ?? '');
	const flipV = /^(1|true)$/.test(xfrm.getAttribute('flipV') ?? '');
	const result: DrawingTransform = {
		...(Number.isFinite(rotation) && rotation ? { rotation } : {}),
		...(flipH ? { flipH } : {}),
		...(flipV ? { flipV } : {}),
	};
	return Object.keys(result).length ? result : undefined;
}

export function writeDrawingTransform(xfrm: XmlElement, transform: DrawingTransform): void {
	const values = {
		rot: transform.rotation ? String(Math.round(transform.rotation * 60000)) : '',
		flipH: transform.flipH ? '1' : '',
		flipV: transform.flipV ? '1' : '',
	};
	for (const [key, value] of Object.entries(values)) {
		if (value) xfrm.setAttribute(key, value);
		else xfrm.removeAttribute(key);
	}
}
