import { parseDrawingColor, resolveDrawingColor, type DrawingColor } from '../drawingml/index';
import { elements } from '../xml/index';

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

/** A saved DrawingML colour element, or one the shared `drawingml` reader already parsed. */
export type ThemeColorSource = Element | DrawingColor;

const isParsed = (source: ThemeColorSource): source is DrawingColor =>
	typeof (source as Partial<DrawingColor>).kind === 'string';

/** Visio's namespace guard over the shared colour reader: the colour and its transforms are `a:`. */
function visioColor(source: ThemeColorSource | undefined): DrawingColor | undefined {
	if (!source) return undefined;
	if (isParsed(source)) return source;
	if (source.namespaceURI !== DRAWING_NS) return undefined;
	if (elements(source).some((transform) => transform.namespaceURI !== DRAWING_NS)) return undefined;
	return parseDrawingColor(source);
}
/** The single `a:` colour of `node` (with only `a:` transforms), read by the shared reader. */
export function parsedColorChoice(node: Element | undefined): DrawingColor | undefined {
	return visioColor(colorChoice(node));
}

// MS-VSDX 2.3.4.2.22 declares hueMod unused in dynamic themes; the rest stay unresolved.
const SUPPORTED_TRANSFORMS = new Set(['shade', 'tint', 'alpha']);

/**
 * Bounded DrawingML color choices; unsupported transforms never silently disappear. The colour is
 * read and its tint, shade and alpha applied (linear light, document order) by `drawingml`; this
 * guard keeps Visio's limits: srgb, saved system and scheme colours only, at most 32 transforms,
 * shade/tint/alpha as integers up to 100000, and scheme references at most 8 deep.
 */
export function drawingPaint(
	node: ThemeColorSource | undefined,
	colors: ReadonlyMap<string, ThemeColorSource>,
	placeholder?: string,
	depth = 0,
): { color: string; opacity: number } | undefined {
	if (depth > 8) return undefined;
	const color = visioColor(node);
	if (!color || color.transforms.length > 32) return undefined;
	const transforms = color.transforms.filter((transform) => transform.name !== 'hueMod');
	if (
		transforms.some(
			(transform) =>
				!SUPPORTED_TRANSFORMS.has(transform.name) ||
				(integer(transform.value) ?? 100_001) > 100_000,
		)
	)
		return undefined;
	let rgb: string | undefined;
	let inherited: { color: string; opacity: number } | undefined;
	if (color.kind === 'srgb') rgb = color.value;
	else if (color.kind === 'system') rgb = color.fallback;
	else if (color.kind === 'scheme') {
		if (color.value === 'phClr') rgb = placeholder;
		else {
			inherited = drawingPaint(colors.get(color.value), colors, placeholder, depth + 1);
			rgb = inherited?.color;
		}
	}
	if (!rgb || !/^#?[\da-f]{6}$/i.test(rgb)) return undefined;
	const resolved = resolveDrawingColor(
		{ kind: 'srgb', value: rgb.replace(/^#/, ''), transforms },
		undefined,
		{ transformOrder: 'document' },
	);
	if (!resolved) return undefined;
	return {
		opacity: transforms.some((transform) => transform.name === 'alpha')
			? resolved.alpha
			: (inherited?.opacity ?? 1),
		color: resolved.hex.toLowerCase(),
	};
}
/** Flat colors cannot represent partial alpha, so leave those unresolved. */
export function drawingColor(
	node: ThemeColorSource | undefined,
	colors: ReadonlyMap<string, ThemeColorSource>,
	placeholder?: string,
): string | undefined {
	const paint = drawingPaint(node, colors, placeholder);
	return paint?.opacity === 1 ? paint.color : undefined;
}
