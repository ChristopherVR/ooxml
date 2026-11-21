import { NS, first, type XmlDocument, type XmlElement } from '../xml/index';
import { parseDrawingTextWarp } from './read-text-body';
import { writeDrawingTextWarp } from './write-text-body';
import type { DrawingTextWarp } from './text-model';

/** Text effects shared by editable WordArt in Word and Excel. Colors are OOXML sRGB. */
export interface WordArtStyle {
	fill: string;
	outline?: string;
	shadow?: string;
	warp?: DrawingTextWarp;
}

export interface WordArtText {
	text: string;
	style: WordArtStyle;
}

export const WORD_ART_FONT = 'Arial';
export const WORD_ART_SIZE_PT = 30;
export const WORD_ART_OUTLINE_EMU = 12700;
export const WORD_ART_SHADOW = { blurRad: 25400, dist: 38100, dir: 2700000, opacity: 55000 };

export const wordArtPaintEquals = (a: WordArtStyle, b: WordArtStyle): boolean =>
	a.fill === b.fill && a.outline === b.outline && a.shadow === b.shadow;

export function wordArtRgb(color: unknown): string | undefined {
	return typeof color === 'string' && /^[0-9a-f]{6}$/i.test(color)
		? color.toUpperCase()
		: undefined;
}

export function readWordArtWarp(body: XmlElement | undefined): DrawingTextWarp | undefined {
	const warp = parseDrawingTextWarp(first(body, 'prstTxWarp', NS.a));
	return warp?.preset === 'textNoShape' ? undefined : warp;
}

/** Changes only the preset and its guides, leaving body layout and other effects in place. */
export function applyWordArtWarp(body: XmlElement, warp: DrawingTextWarp | undefined): void {
	const previous = first(body, 'prstTxWarp', NS.a);
	if (JSON.stringify(readWordArtWarp(body)) === JSON.stringify(warp)) return;
	if (previous) body.removeChild(previous);
	if (!warp || !body.ownerDocument) return;
	body.insertBefore(writeDrawingTextWarp(body.ownerDocument, warp), body.firstChild);
}

export function wordArtShadowElement(
	doc: XmlDocument,
	color: string,
	namespace: string,
	prefix: 'a' | 'w14',
): XmlElement {
	const shadow = doc.createElementNS(namespace, prefix === 'a' ? 'a:outerShdw' : 'w14:shadow');
	const rgb = doc.createElementNS(namespace, `${prefix}:srgbClr`);
	const alpha = doc.createElementNS(namespace, `${prefix}:alpha`);
	const attr = (node: XmlElement, name: string, value: string | number) => {
		if (prefix === 'a') node.setAttribute(name, String(value));
		else node.setAttributeNS(namespace, `w14:${name}`, String(value));
	};
	const { blurRad, dist, dir, opacity } = WORD_ART_SHADOW;
	attr(shadow, 'blurRad', blurRad);
	attr(shadow, 'dist', dist);
	attr(shadow, 'dir', dir);
	attr(rgb, 'val', color);
	attr(alpha, 'val', opacity);
	rgb.appendChild(alpha);
	shadow.appendChild(rgb);
	return shadow;
}
