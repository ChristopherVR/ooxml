import { linearToSrgb255, srgb255ToLinear } from '../color/index.js';
import { elements } from '../xml/index.js';

export const DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const THEME_NS = 'http://schemas.microsoft.com/office/visio/2012/theme';
export const themeChildren = (node: Element | undefined, name: string, ns = DRAWING_NS) =>
	(node ? elements(node) : []).filter(
		(item) => item.namespaceURI === ns && item.localName === name,
	);
export const themeChild = (node: Element | undefined, name: string, ns = DRAWING_NS) =>
	themeChildren(node, name, ns)[0];
export function extension(node: Element | undefined, name: string): Element | undefined {
	return themeChildren(themeChild(node, 'extLst'), 'ext').flatMap((ext) =>
		themeChildren(ext, name, THEME_NS),
	)[0];
}
export function integer(value: string | null | undefined): number | undefined {
	if (!value || !/^\d{1,6}$/.test(value)) return undefined;
	return Number(value);
}
export function colorChoice(node: Element | undefined): Element | undefined {
	const choices = (node ? elements(node) : []).filter((item) => item.namespaceURI === DRAWING_NS);
	return choices.length === 1 ? choices[0] : undefined;
}
/** Bounded DrawingML color choices; unsupported transforms never silently disappear. */
export function drawingPaint(
	node: Element | undefined,
	colors: ReadonlyMap<string, Element>,
	placeholder?: string,
	depth = 0,
): { color: string; opacity: number } | undefined {
	if (!node || depth > 8 || node.namespaceURI !== DRAWING_NS) return undefined;
	let rgb: string | undefined;
	let opacity = 1;
	if (node.localName === 'srgbClr') rgb = node.getAttribute('val') ?? undefined;
	else if (node.localName === 'sysClr') rgb = node.getAttribute('lastClr') ?? undefined;
	else if (node.localName === 'schemeClr') {
		const name = node.getAttribute('val') ?? '';
		const inherited =
			name === 'phClr' ? undefined : drawingPaint(colors.get(name), colors, placeholder, depth + 1);
		rgb = name === 'phClr' ? placeholder : inherited?.color;
		opacity = inherited?.opacity ?? 1;
	}
	if (!rgb || !/^#?[\da-f]{6}$/i.test(rgb)) return undefined;
	rgb = rgb.replace(/^#/, '');
	let channels = [0, 2, 4].map((offset) =>
		srgb255ToLinear(parseInt(rgb.slice(offset, offset + 2), 16)),
	);
	const transforms = elements(node);
	if (transforms.length > 32) return undefined;
	for (const transform of transforms) {
		// MS-VSDX 2.3.4.2.22 declares hueMod unused in dynamic themes.
		if (transform.namespaceURI === DRAWING_NS && transform.localName === 'hueMod') continue;
		const value = integer(transform.getAttribute('val'));
		if (transform.namespaceURI !== DRAWING_NS || value === undefined || value > 100_000)
			return undefined;
		const ratio = value / 100_000;
		// ISO 29500 tint/shade mix with white/black in linear light.
		if (transform.localName === 'shade') channels = channels.map((channel) => channel * ratio);
		else if (transform.localName === 'tint')
			channels = channels.map((channel) => channel * ratio + 1 - ratio);
		else if (transform.localName === 'alpha') opacity = ratio;
		else return undefined;
	}
	return {
		opacity,
		color: `#${channels.map((channel) => Math.round(linearToSrgb255(channel)).toString(16).padStart(2, '0')).join('')}`,
	};
}
/** Flat colors cannot represent partial alpha, so leave those unresolved. */
export function drawingColor(
	node: Element | undefined,
	colors: ReadonlyMap<string, Element>,
	placeholder?: string,
): string | undefined {
	const paint = drawingPaint(node, colors, placeholder);
	return paint?.opacity === 1 ? paint.color : undefined;
}
