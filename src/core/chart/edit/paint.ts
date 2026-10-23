// Colour and fill primitives every chart edit shares (lifted from the xlsx edit layer, see
// PROVENANCE.md): DrawingML colour equality including the ordered transforms, transparency as an
// `a:alpha` transform that keeps theme and luminance transforms, and recolouring a gradient whose
// stops follow a palette colour. Pure functions over `drawingml` types; inputs are never mutated.
import { resolveDrawingColor } from '../../drawingml/drawing-color';
import type { DrawingColor, DrawingFill } from '../../drawingml/types';

const ALPHA = new Set(['alpha', 'alphaMod', 'alphaOff']);

/** Equality of DrawingML colours, including the ordered transform sequence. */
export function sameDrawingColor(
	a: DrawingColor | undefined,
	b: DrawingColor | undefined,
): boolean {
	if (!a || !b) return a === b;
	return (
		a.kind === b.kind &&
		a.value === b.value &&
		a.fallback === b.fallback &&
		a.transforms.length === b.transforms.length &&
		a.transforms.every(
			(t, i) => b.transforms[i]?.name === t.name && b.transforms[i]?.value === t.value,
		)
	);
}

/** The opacity transforms (`alpha`, `alphaMod`, `alphaOff`) of a colour, cloned. */
export const drawingColorAlphaTransforms = (
	color: DrawingColor | undefined,
): DrawingColor['transforms'] =>
	structuredClone(color?.transforms.filter((item) => ALPHA.has(item.name)) ?? []);

/** `color` with the opacity transforms of `from` appended (a recolour keeps transparency). */
export function withAlphaOf(color: DrawingColor, from: DrawingColor | undefined): DrawingColor {
	const next = structuredClone(color);
	next.transforms.push(...drawingColorAlphaTransforms(from));
	return next;
}

/** Transparency in percent (0..100) after the ordered transforms; the scheme is irrelevant. */
export function drawingColorTransparency(color: DrawingColor): number | undefined {
	const resolved = resolveDrawingColor(
		color,
		{ scheme: () => '#000000' },
		{ transformOrder: 'document' },
	);
	return resolved ? Math.round((1 - resolved.alpha) * 100000) / 1000 : undefined;
}

/** Checks a 0..100 percentage, with the message the chart editors show. */
export function assertChartPercent(value: number, message: string): void {
	if (!Number.isFinite(value) || value < 0 || value > 100) throw new RangeError(message);
}

/** `color` with its opacity transforms replaced by one `a:alpha` for `percent` transparency. */
export function withDrawingColorTransparency(color: DrawingColor, percent: number): DrawingColor {
	assertChartPercent(percent, 'Chart transparency must be from 0 to 100');
	const next = structuredClone(color);
	next.transforms = next.transforms.filter((item) => !ALPHA.has(item.name));
	if (percent !== 0)
		next.transforms.push({ name: 'alpha', value: String(Math.round((100 - percent) * 1000)) });
	return next;
}

/**
 * A gradient whose every stop is `previous` (plus its own shading transforms) recoloured to `next`;
 * any other fill comes back unchanged. Palette transforms precede gradient shading transforms in
 * the XML Office saves, so the stop keeps the transforms after the palette's.
 */
export function recolorPaletteGradient(
	fill: DrawingFill | undefined,
	previous: DrawingColor,
	next: DrawingColor,
): DrawingFill | undefined {
	if (fill?.kind !== 'gradient') return fill;
	const matches = fill.stops.every((stop) =>
		sameDrawingColor(
			{ ...stop.color, transforms: stop.color.transforms.slice(0, previous.transforms.length) },
			previous,
		),
	);
	if (!matches) return fill;
	return {
		...fill,
		stops: fill.stops.map((stop) => ({
			...stop,
			color: {
				...next,
				transforms: [
					...next.transforms,
					...stop.color.transforms.slice(previous.transforms.length),
				],
			},
		})),
	};
}
