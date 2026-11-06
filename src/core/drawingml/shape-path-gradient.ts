/**
 * `a:gradFill/a:path@path="shape"`: what PowerPoint paints, measured on
 * PowerPoint 16 (`Slide.Export`, white to black, inverted through the sigma
 * curve to a position field and compared with candidate fields on 130 presets
 * and two freeforms).
 *
 * The field is not a distance to the outline. Position 0 sits at the focus
 * (the centre of `a:fillToRect`) and position 1 on the outline, along every
 * ray from the focus, so the iso-lines are the outline scaled about the focus
 * (`t = |p - f| / |b(p) - f|`, `b(p)` where the ray leaves the shape). The
 * mean absolute position error against PowerPoint is 0.002 to 0.008 for the
 * ellipse, triangles, stars, rounded rectangles, hexagon and the like, with
 * the focus at the centre or off it (even outside the shape).
 *
 * Two families do not follow it: every block arrow, the chevron, the swoosh,
 * the lightning bolt, sun, moon, ribbons, scrolls, waves, a few flow chart
 * symbols and the smiley face, plus any freeform (`a:custGeom`), paint a true
 * circle to the farthest corner of the box, exactly as `path="circle"` does
 * (error 0.001 to 0.005). The ring-shaped presets (donut, frame, block arc, ...)
 * match neither and keep their approximation.
 *
 * The scaled outline is drawn as nested copies of the outline, largest first,
 * each filled with the colour at the middle of its band, so the markup is plain
 * SVG that renders in any `background-image`, raster export or older browser.
 * A stop with alpha is drawn as rings (an outer copy with the next copy cut out
 * by the even-odd rule) so the layers do not stack.
 *
 * @module shape-path-gradient
 */
import { gradientColorAt, type RectPathGradientStop } from './rect-path-gradient';

/** How a `path="shape"` gradient is painted on a given geometry. */
export type ShapePathField = 'outline' | 'circle' | 'legacy';

/** Presets measured to paint `path="shape"` as the circle to the farthest corner. */
const CIRCLE_FIELD_PRESETS = new Set(
	[
		'chevron',
		'rightArrow',
		'leftArrow',
		'upArrow',
		'downArrow',
		'stripedRightArrow',
		'notchedRightArrow',
		'bentUpArrow',
		'bentArrow',
		'uturnArrow',
		'leftRightArrow',
		'upDownArrow',
		'leftUpArrow',
		'leftRightUpArrow',
		'quadArrow',
		'curvedRightArrow',
		'curvedLeftArrow',
		'curvedUpArrow',
		'curvedDownArrow',
		'swooshArrow',
		'lightningBolt',
		'sun',
		'moon',
		'smileyFace',
		'ribbon',
		'ribbon2',
		'ellipseRibbon',
		'ellipseRibbon2',
		'leftRightRibbon',
		'verticalScroll',
		'horizontalScroll',
		'wave',
		'doubleWave',
		'flowChartDocument',
		'flowChartMultidocument',
		'flowChartPunchedTape',
		'flowChartMagneticDisk',
		'flowChartDisplay',
	].map((name) => name.toLowerCase()),
);

/** Presets whose outline has holes or detached parts: no measured field, keep the approximation. */
const LEGACY_FIELD_PRESETS = new Set(
	[
		'donut',
		'noSmoking',
		'blockArc',
		'frame',
		'chord',
		'funnel',
		'mathDivide',
		'mathNotEqual',
		'cornerTabs',
		'squareTabs',
		'plaqueTabs',
		'circularArrow',
		'leftCircularArrow',
		'leftRightCircularArrow',
		'flowChartMagneticTape',
		'flowChartMagneticDrum',
	].map((name) => name.toLowerCase()),
);

/**
 * Which field a `path="shape"` gradient paints on `preset`. Pass `undefined`
 * for a freeform (`a:custGeom`), which always paints the circle.
 */
export function shapePathGradientField(preset: string | undefined): ShapePathField {
	if (!preset || preset === 'custom') return 'circle';
	const key = preset.toLowerCase();
	if (CIRCLE_FIELD_PRESETS.has(key)) return 'circle';
	if (LEGACY_FIELD_PRESETS.has(key)) return 'legacy';
	return 'outline';
}

/** Bands across the 0 to 1 position range (each step is a hair under 1% of the ramp). */
const BAND_COUNT = 100;

function n(value: number): string {
	return String(Math.round(value * 1000) / 1000);
}

const ARG_COUNT: Record<string, number> = { M: 2, L: 2, Q: 4, C: 6, A: 7, Z: 0 };

/**
 * Scales an absolute SVG path (`M L Q C A Z`) about `(fx, fy)`. Returns
 * `undefined` for any other command, so the caller can fall back.
 */
export function scalePathAbout(d: string, s: number, fx: number, fy: number): string | undefined {
	const tokens = d.match(/[A-Za-z]|-?\d*\.?\d+(?:e[-+]?\d+)?/gu);
	if (!tokens) return undefined;
	const out: string[] = [];
	let i = 0;
	while (i < tokens.length) {
		const command = tokens[i++]!;
		const count = ARG_COUNT[command];
		if (count === undefined) return undefined;
		out.push(command);
		if (command === 'Z') continue;
		// An implicit repeat of M is L; the evaluator never writes one, so
		// reading exactly one group per command is enough.
		const args = tokens.slice(i, i + count).map(Number);
		i += count;
		if (args.length !== count || args.some((value) => !Number.isFinite(value))) return undefined;
		if (command === 'A') {
			out.push(
				n(args[0]! * s),
				n(args[1]! * s),
				n(args[2]!),
				n(args[3]!),
				n(args[4]!),
				n(fx + (args[5]! - fx) * s),
				n(fy + (args[6]! - fy) * s),
			);
			continue;
		}
		for (let k = 0; k < count; k += 2) {
			out.push(n(fx + (args[k]! - fx) * s), n(fy + (args[k + 1]! - fy) * s));
		}
	}
	return out.join(' ');
}

/**
 * The closed subpaths of `d` (the evaluator appends open stroke-only lines,
 * such as a flow chart's inner bars, to the same path), joined, or `undefined`
 * when there is none.
 */
export function closedSubpaths(d: string): string | undefined {
	const parts = d
		.split(/(?=M)/u)
		.map((part) => part.trim())
		.filter((part) => part.endsWith('Z'));
	return parts.length > 0 ? parts.join(' ') : undefined;
}

/**
 * The scaled-outline field as a self-contained SVG (viewBox in the shape's own
 * pixels, `preserveAspectRatio="none"`): `d` is the outline in a
 * `width` x `height` box and `(focusX, focusY)` the focus as fractions of it.
 * Returns `undefined` when the outline cannot be scaled (an unsupported path
 * command or no stops).
 */
export function buildOutlinePathGradientSvg(
	stops: readonly RectPathGradientStop[],
	d: string,
	width: number,
	height: number,
	focusX: number,
	focusY: number,
): string | undefined {
	if (stops.length === 0 || !(width > 0) || !(height > 0)) return undefined;
	if (scalePathAbout(d, 1, 0, 0) === undefined) return undefined;
	const sorted = [...stops].sort((a, b) => a.position - b.position);
	const fx = focusX * width;
	const fy = focusY * height;
	const translucent = sorted.some((stop) => (stop.opacity ?? 1) < 1);
	const body: string[] = [];
	if (translucent) {
		for (let i = BAND_COUNT; i >= 1; i -= 1) {
			const outer = scalePathAbout(d, i / BAND_COUNT, fx, fy);
			if (outer === undefined) return undefined;
			const inner = i > 1 ? scalePathAbout(d, (i - 1) / BAND_COUNT, fx, fy) : undefined;
			const { color, opacity } = gradientColorAt(sorted, ((i - 0.5) / BAND_COUNT) * 100);
			const alpha = opacity < 1 ? ` fill-opacity="${n(Math.max(0, opacity))}"` : '';
			body.push(
				`<path d="${outer}${inner ? ` ${inner}` : ''}" fill="${color}"${alpha} fill-rule="evenodd"/>`,
			);
		}
		return wrap(width, height, body.join(''));
	}
	body.push(`<defs><path id="p" d="${d}"/></defs>`);
	for (let i = BAND_COUNT; i >= 1; i -= 1) {
		const s = i / BAND_COUNT;
		const { color } = gradientColorAt(sorted, ((i - 0.5) / BAND_COUNT) * 100);
		body.push(
			`<use href="#p" transform="matrix(${n(s)} 0 0 ${n(s)} ${n(fx * (1 - s))} ${n(fy * (1 - s))})" fill="${color}"/>`,
		);
	}
	return wrap(width, height, body.join(''));
}

function wrap(width: number, height: number, body: string): string {
	return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${n(width)} ${n(height)}" preserveAspectRatio="none">${body}</svg>`;
}
