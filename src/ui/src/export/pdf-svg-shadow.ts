/**
 * Soft shadows for the SVG to PDF converter: a Gaussian blur filter on a shadow group (Visio's
 * `data-shadow`, a chart format shadow) and an `feDropShadow` filter on a shape (chart series)
 * are painted as a blurred silhouette (`pdf-blur-mask.ts`) instead of being dropped or drawn
 * hard-edged. Only geometry is blurred: a subtree holding text or pictures is left to the
 * caller, which reports it.
 */
import { drawBlurredSilhouette, type SilhouettePart } from './pdf-blur-mask';
import { parsePath } from './pdf-path';
import { childElements } from './pdf-svg-dom';
import { basicShapePath } from './pdf-svg-shapes';
import { number, parseColor, parseTransform, PdfFallback, type Rgba } from './pdf-svg-style';
import type { PdfVectorPage } from './pdf-vector';

/** What a filter does that the converter can paint. */
export type SvgShadowFilter =
	| { kind: 'blur'; sigma: number }
	| { kind: 'drop'; sigma: number; dx: number; dy: number; color: Rgba };

/** How the element was handled: drawn as its blurred self, given a shadow beneath, or not. */
export type SvgShadowResult = 'replaced' | 'under' | 'none';

const NOT_DRAWN = new Set([
	'defs',
	'title',
	'desc',
	'metadata',
	'clipPath',
	'marker',
	'symbol',
	'linearGradient',
	'radialGradient',
	'filter',
	'mask',
	'style',
]);

const deviation = (value: string | null): number => {
	const values = (value ?? '')
		.trim()
		.split(/[\s,]+/)
		.map(Number)
		.filter((v) => Number.isFinite(v) && v >= 0);
	return values.length ? Math.max(...values) : 0;
};

/** The filter as a blur or a drop shadow, or `undefined` for anything else. */
export function readShadowFilter(filter: Element | undefined): SvgShadowFilter | undefined {
	if (filter?.localName !== 'filter') return undefined;
	const primitives = childElements(filter);
	const only = primitives.length === 1 ? primitives[0] : undefined;
	if (only?.localName === 'feGaussianBlur') {
		const input = only.getAttribute('in');
		if (input && input !== 'SourceGraphic') return undefined;
		const sigma = deviation(only.getAttribute('stdDeviation'));
		return sigma > 0 ? { kind: 'blur', sigma } : undefined;
	}
	if (only?.localName === 'feDropShadow') {
		let color: Rgba | null;
		try {
			color = parseColor(only.getAttribute('flood-color') ?? '#000000');
		} catch {
			return undefined;
		}
		if (!color) return undefined;
		return {
			kind: 'drop',
			sigma: deviation(only.getAttribute('stdDeviation') ?? '2'),
			dx: number(only.getAttribute('dx'), 2),
			dy: number(only.getAttribute('dy'), 2),
			color: { ...color, a: color.a * number(only.getAttribute('flood-opacity'), 1) },
		};
	}
	return undefined;
}

const compose = (a: readonly number[], b: readonly number[]): number[] => [
	a[0]! * b[0]! + a[2]! * b[1]!,
	a[1]! * b[0]! + a[3]! * b[1]!,
	a[0]! * b[2]! + a[2]! * b[3]!,
	a[1]! * b[2]! + a[3]! * b[3]!,
	a[0]! * b[4]! + a[2]! * b[5]! + a[4]!,
	a[1]! * b[4]! + a[3]! * b[5]! + a[5]!,
];

interface Collected {
	parts: SilhouettePart[];
	/** The flat paints the parts use (with their alpha). */
	colors: Rgba[];
}

/** The shapes an element draws, in its own space; `false` when it draws text or pictures. */
function collect(
	element: Element,
	matrix: readonly number[],
	into: Collected,
	alpha: number,
	depth = 0,
): boolean {
	const name = element.localName;
	if (NOT_DRAWN.has(name) || element.getAttribute('display') === 'none') return true;
	if (depth > 64) return false;
	// The root's own opacity is in the caller's `opacity` already.
	const own = depth ? alpha * number(element.getAttribute('opacity'), 1) : alpha;
	if (name === 'g' || name === 'svg') {
		for (const child of childElements(element)) {
			const transforms = parseTransform(child.getAttribute('transform'));
			const childMatrix = transforms.reduce<readonly number[]>(compose, matrix);
			if (!collect(child, childMatrix, into, own, depth + 1)) return false;
		}
		return true;
	}
	let d: string | undefined;
	if (name === 'path') d = element.getAttribute('d') ?? '';
	else if (name === 'rect') {
		const [x, y, w, h] = ['x', 'y', 'width', 'height'].map((a) => number(element.getAttribute(a)));
		if (w! > 0 && h! > 0) d = `M ${x} ${y} h ${w} v ${h} h ${-w!} Z`;
		else return true;
	} else if (['circle', 'ellipse', 'line', 'polyline', 'polygon'].includes(name))
		d = basicShapePath(element);
	else return false;
	const segments = d === undefined ? [] : parsePath(d);
	if (!segments) throw new PdfFallback('path data');
	if (!segments.length) return true;
	const paint = (value: string | null): Rgba | 'other' | null => {
		if (value?.trim().startsWith('url(')) return 'other';
		const color = parseColor(value);
		return color && { ...color, a: color.a * own };
	};
	const fill = name === 'line' ? null : paint(element.getAttribute('fill') ?? '#000000');
	const stroke = paint(element.getAttribute('stroke'));
	const width = number(element.getAttribute('stroke-width'), 1);
	const part: SilhouettePart = { segments, matrix };
	if (fill) part.fill = element.getAttribute('fill-rule') === 'evenodd' ? 'evenodd' : 'nonzero';
	if (stroke && width > 0) part.stroke = width;
	if (!part.fill && !part.stroke) return true;
	into.parts.push(part);
	for (const used of [part.fill ? fill : null, part.stroke ? stroke : null])
		if (used && used !== 'other') into.colors.push(used);
	return true;
}

const sameColor = (a: Rgba, b: Rgba) =>
	Math.abs(a.r - b.r) + Math.abs(a.g - b.g) + Math.abs(a.b - b.b) + Math.abs(a.a - b.a) < 1e-3;

/**
 * Paints `element`'s soft shadow for `filter` in the element's own space (its transform already
 * applied). A blur of a one-colour shadow group replaces the group; a drop shadow is painted
 * beneath the element, which the caller then draws as usual.
 */
export function paintSvgShadow(
	page: PdfVectorPage,
	element: Element,
	filter: SvgShadowFilter,
	opacity: number,
): SvgShadowResult {
	const collected: Collected = { parts: [], colors: [] };
	const identity = [1, 0, 0, 1, 0, 0];
	try {
		if (!collect(element, identity, collected, 1)) return 'none';
	} catch (error) {
		// What the shadow cannot read, the caller's own drawing reports.
		if (error instanceof PdfFallback) return 'none';
		throw error;
	}
	if (!collected.parts.length) return filter.kind === 'blur' ? 'replaced' : 'none';
	if (filter.kind === 'blur') {
		const [color] = collected.colors;
		if (!color || collected.colors.some((other) => !sameColor(other, color))) return 'none';
		drawBlurredSilhouette(page, collected.parts, filter.sigma, color, color.a * opacity);
		return 'replaced';
	}
	const shifted = collected.parts.map((part) => ({
		...part,
		matrix: compose([1, 0, 0, 1, filter.dx, filter.dy], part.matrix ?? identity),
	}));
	drawBlurredSilhouette(page, shifted, filter.sigma, filter.color, filter.color.a * opacity);
	return 'under';
}
