import { NS, elements, first, type XmlDocument, type XmlElement } from '../xml/index';
import { parseDrawingTextWarp } from './read-text-body';
import { writeDrawingTextWarp } from './write-text-body';
import type { DrawingTextWarp } from './text-model';

/** Glow effect on WordArt text (`w14:glow` / `a:glow`). */
export interface WordArtGlow {
	/** OOXML sRGB hex (`RRGGBB`). */
	color: string;
	/** Glow radius in EMU (`rad`). */
	radiusEmu: number;
	/** Opacity in OOXML 1/100000 units (default 40000). */
	opacity?: number;
}

/** Reflection effect on WordArt text (`w14:reflection` / `a:reflection`). */
export interface WordArtReflection {
	/** Blur radius in EMU (`blurRad`). */
	blurRadiusEmu?: number;
	/** Starting opacity in 1/100000 units (`stA`, default 50000). */
	startOpacity?: number;
	/** Ending opacity in 1/100000 units (`endA`, default 0). */
	endOpacity?: number;
	/** Distance below text in EMU (`dist`, default 0). */
	distanceEmu?: number;
	/** Vertical scale percentage in 1/100000 units (`sy`, default -100000). */
	scaleY?: number;
}

/** Text effects shared by editable WordArt in Word and Excel. Colors are OOXML sRGB. */
export interface WordArtStyle {
	fill: string;
	gradient?: WordArtGradient;
	outline?: string;
	shadow?: string;
	glow?: WordArtGlow;
	reflection?: WordArtReflection;
	warp?: DrawingTextWarp;
}

/** Two-stop sRGB gradient with a DrawingML clockwise angle in 1/60000 degrees. */
export interface WordArtGradient {
	end: string;
	angle: number;
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
	a.fill === b.fill &&
	a.outline === b.outline &&
	a.shadow === b.shadow &&
	a.gradient?.end === b.gradient?.end &&
	a.gradient?.angle === b.gradient?.angle &&
	JSON.stringify(a.glow) === JSON.stringify(b.glow) &&
	JSON.stringify(a.reflection) === JSON.stringify(b.reflection);

export function wordArtRgb(color: unknown): string | undefined {
	return typeof color === 'string' && /^[0-9a-f]{6}$/i.test(color)
		? color.toUpperCase()
		: undefined;
}

export function normalizeWordArtGradient(
	gradient: WordArtGradient | undefined,
): WordArtGradient | undefined {
	const end = wordArtRgb(gradient?.end);
	const angle = gradient?.angle;
	return end &&
		typeof angle === 'number' &&
		Number.isInteger(angle) &&
		angle >= 0 &&
		angle < 21600000
		? { end, angle }
		: undefined;
}

/** Reads solid and two-stop linear sRGB text fills in DrawingML or Word 2010. */
export function readWordArtFill(
	parent: XmlElement | undefined,
	namespace: string,
): Pick<WordArtStyle, 'fill' | 'gradient'> | undefined {
	const attr = (node: XmlElement | undefined, name: string) =>
		namespace === NS.a ? node?.getAttribute(name) : node?.getAttributeNS(namespace, name);
	const rgb = (node: XmlElement | undefined) => {
		const color = first(node, 'srgbClr', namespace);
		return color && elements(color).length === 0 ? wordArtRgb(attr(color, 'val')) : undefined;
	};
	const gradient = first(parent, 'gradFill', namespace);
	if (!gradient) {
		const fill = rgb(first(parent, 'solidFill', namespace));
		return fill ? { fill } : undefined;
	}
	const list = first(gradient, 'gsLst', namespace);
	const stops = list ? elements(list) : [];
	const [firstStop, lastStop] = stops;
	if (stops.length !== 2 || attr(firstStop, 'pos') !== '0' || attr(lastStop, 'pos') !== '100000')
		return undefined;
	const fill = rgb(firstStop);
	const end = rgb(lastStop);
	const direction = attr(first(gradient, 'lin', namespace), 'ang');
	const paint =
		direction && end ? normalizeWordArtGradient({ end, angle: Number(direction) }) : undefined;
	return fill && paint ? { fill, gradient: paint } : undefined;
}

export function wordArtFillElement(
	doc: XmlDocument,
	style: Pick<WordArtStyle, 'fill' | 'gradient'>,
	namespace: string,
	prefix: 'a' | 'w14',
): XmlElement {
	const make = (name: string, attrs: Record<string, string> = {}) => {
		const node = doc.createElementNS(namespace, `${prefix}:${name}`);
		for (const [key, value] of Object.entries(attrs)) {
			if (prefix === 'a') node.setAttribute(key, value);
			else node.setAttributeNS(namespace, `${prefix}:${key}`, value);
		}
		return node;
	};
	const fill = wordArtRgb(style.fill) ?? '26374A';
	const gradient = normalizeWordArtGradient(style.gradient);
	if (!gradient) {
		const solid = make('solidFill');
		solid.appendChild(make('srgbClr', { val: fill }));
		return solid;
	}
	const result = make('gradFill');
	const list = make('gsLst');
	for (const [pos, val] of [
		['0', fill],
		['100000', gradient.end],
	] as const) {
		const stop = make('gs', { pos });
		stop.appendChild(make('srgbClr', { val }));
		list.appendChild(stop);
	}
	result.appendChild(list);
	result.appendChild(make('lin', { ang: String(gradient.angle), scaled: '0' }));
	return result;
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

/** Reads a `w14:glow` or `a:glow` element. */
export function readWordArtGlow(
	node: XmlElement | undefined,
	namespace: string,
): WordArtGlow | undefined {
	if (!node) return undefined;
	const attr = (el: XmlElement | undefined, name: string) =>
		namespace === NS.a ? el?.getAttribute(name) : el?.getAttributeNS(namespace, name);
	const rad = Number(attr(node, 'rad'));
	const rgbNode = first(node, 'srgbClr', namespace);
	const color = wordArtRgb(attr(rgbNode, 'val'));
	if (!color || !Number.isFinite(rad) || rad <= 0) return undefined;
	const alphaVal = Number(attr(first(rgbNode, 'alpha', namespace), 'val'));
	return {
		color,
		radiusEmu: rad,
		...(Number.isFinite(alphaVal) && alphaVal >= 0 ? { opacity: alphaVal } : {}),
	};
}

/** Reads a `w14:reflection` or `a:reflection` element. */
export function readWordArtReflection(
	node: XmlElement | undefined,
	namespace: string,
): WordArtReflection | undefined {
	if (!node) return undefined;
	const num = (name: string) => {
		const raw = namespace === NS.a ? node.getAttribute(name) : node.getAttributeNS(namespace, name);
		if (raw === null || raw === '') return undefined;
		const parsed = Number(raw);
		return Number.isFinite(parsed) ? parsed : undefined;
	};
	const blurRadiusEmu = num('blurRad');
	const startOpacity = num('stA');
	const endOpacity = num('endA');
	const distanceEmu = num('dist');
	const scaleY = num('sy');
	return {
		...(blurRadiusEmu !== undefined ? { blurRadiusEmu } : {}),
		...(startOpacity !== undefined ? { startOpacity } : {}),
		...(endOpacity !== undefined ? { endOpacity } : {}),
		...(distanceEmu !== undefined ? { distanceEmu } : {}),
		...(scaleY !== undefined ? { scaleY } : {}),
	};
}

export function wordArtGlowElement(
	doc: XmlDocument,
	glow: WordArtGlow,
	namespace: string,
	prefix: 'a' | 'w14',
): XmlElement {
	const node = doc.createElementNS(namespace, `${prefix}:glow`);
	const rgb = doc.createElementNS(namespace, `${prefix}:srgbClr`);
	const attr = (el: XmlElement, name: string, value: string | number) => {
		if (prefix === 'a') el.setAttribute(name, String(value));
		else el.setAttributeNS(namespace, `w14:${name}`, String(value));
	};
	attr(node, 'rad', Math.round(glow.radiusEmu));
	attr(rgb, 'val', wordArtRgb(glow.color) ?? '000000');
	if (glow.opacity !== undefined) {
		const alpha = doc.createElementNS(namespace, `${prefix}:alpha`);
		attr(alpha, 'val', Math.round(glow.opacity));
		rgb.appendChild(alpha);
	}
	node.appendChild(rgb);
	return node;
}

export function wordArtReflectionElement(
	doc: XmlDocument,
	reflection: WordArtReflection,
	namespace: string,
	prefix: 'a' | 'w14',
): XmlElement {
	const node = doc.createElementNS(namespace, `${prefix}:reflection`);
	const attr = (name: string, value: number | undefined) => {
		if (value === undefined) return;
		if (prefix === 'a') node.setAttribute(name, String(Math.round(value)));
		else node.setAttributeNS(namespace, `w14:${name}`, String(Math.round(value)));
	};
	attr('blurRad', reflection.blurRadiusEmu);
	attr('stA', reflection.startOpacity);
	attr('endA', reflection.endOpacity);
	attr('dist', reflection.distanceEmu);
	attr('sy', reflection.scaleY);
	return node;
}
