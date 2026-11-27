import type { InkStroke, InkXmlObject } from './types';
import {
	countAuthoredInkPathSamples,
	countInkTraceSamples,
	indexInkMl,
	MAX_INK_PART_SAMPLES,
} from './inkml-limits';
import type { InkTargetBox } from './inkml-ink-space';
import { inkBounds, inkLengthToPx, inkPointMapper } from './inkml-ink-space';
import { inkMlTiming } from './inkml-timing';
import {
	decodeTracePoints,
	ensureArray,
	nsAttr,
	nsGet,
	pointsToPressures,
	pointsToSvgPath,
	pointsToTilt,
	pointsToTimestamps,
	resolveChannelOrder,
} from './inkml-trace-decode';
import { isFiniteNumber } from '../../units/finite';

// Re-exported for the existing import sites (and colocated tests): the
// writer half of this module was split out to `inkml-content-part-writer.ts`
// to keep both files under this repo's file-size guideline.
export { buildInkMlContent } from './inkml-content-part-writer';

export interface ParsedInkMlContent {
	strokes: InkStroke[];
	rawXml: InkXmlObject;
	/**
	 * Part of the ink was not read: a trace over the sample limit, or the part over its sample or
	 * trace budget. Rewriting the part from the strokes read would lose the rest, so callers keep
	 * such a part as it is and do not edit its strokes.
	 */
	truncated: boolean;
}

type BrushStyle = Pick<InkStroke, 'color' | 'width' | 'opacity' | 'tool' | 'brushExtras'>;

const DEFAULT_BRUSH: BrushStyle = { color: '#000000', width: 1, opacity: 1 };

/**
 * Parse authored InkML trace/brush metadata while tolerating plain legacy traces.
 *
 * `box` is the `p:contentPart` extent in CSS pixels. When it is supplied, every
 * trace decoded from raw channel data is normalised into that box (see
 * `inkml-ink-space`), because a real PowerPoint InkML part is written in its own
 * device units and would otherwise land thousands of pixels off-element. Strokes
 * carrying the library's own authored `@pva:path` are already in element space
 * and pass through untouched.
 */
export function parseInkMlContent(data: InkXmlObject, box?: InkTargetBox): ParsedInkMlContent {
	const index = indexInkMl(data);
	const inkRoot = nsGet(data, 'ink');
	if (!index || !inkRoot) return { strokes: [], rawXml: data, truncated: Boolean(inkRoot) };
	let truncated = false;
	const space = authoredSpace(inkRoot);
	const brushes = collectBrushes(index.brushes);
	const timing = inkMlTiming(index);
	const channelOrders = new Map<InkXmlObject | undefined, string[]>();

	// Bounds cover the whole part, so normalisation keeps strokes together.
	let remainingSamples = MAX_INK_PART_SAMPLES;
	const decoded = index.traces.flatMap((trace) => {
		const text = typeof trace === 'object' ? String(trace['#text'] ?? '') : String(trace);
		const authored = typeof trace === 'object' ? String(nsAttr(trace, 'path') ?? '') : '';
		const traceSamples = countInkTraceSamples(text);
		const pathSamples = countAuthoredInkPathSamples(authored);
		if (traceSamples === undefined || pathSamples === undefined) {
			truncated = true;
			return [];
		}
		const count = Math.max(traceSamples, pathSamples);
		if (count > remainingSamples) {
			truncated = true;
			return [];
		}
		remainingSamples -= count;
		const format = timing.formatFor(trace);
		let channelOrder = channelOrders.get(format);
		if (!channelOrder) {
			channelOrder = resolveChannelOrder(format);
			channelOrders.set(format, channelOrder);
		}
		const order = authored
			? [
					'X',
					'Y',
					'F',
					...(channelOrder.includes('AZIMUTH') ? ['AZIMUTH', 'ALTITUDE'] : ['OTX', 'OTY']),
					...(channelOrder.includes('T') ? ['T'] : []),
				]
			: channelOrder;
		// With T but without a tilt pair, time occupies the fourth column.
		if (authored && channelOrder.includes('T') && channelOrder.indexOf('T') === 3) {
			order.splice(3, 2);
		}
		return [
			{
				trace,
				authored: authored.trim(),
				force: channelRange(format, 'F'),
				order,
				points: decodeTracePoints(text, order),
				untimed: typeof trace === 'object' && nsAttr(trace, 'untimed') !== undefined,
				brushRef:
					typeof trace === 'object' ? String(nsAttr(trace, 'brushRef') ?? '').replace('#', '') : '',
			},
		];
	});

	const bounds = box
		? inkBounds(decoded.filter((entry) => !entry.authored).map((entry) => entry.points))
		: undefined;
	const mapPoint = bounds && box ? inkPointMapper(bounds, box) : undefined;

	const strokes: InkStroke[] = [];
	for (const entry of decoded) {
		const points =
			mapPoint && !entry.authored
				? mapDecodedPoints(entry.points, entry.order, mapPoint)
				: entry.points;
		const path =
			(entry.authored && scaleAuthoredPath(entry.authored, space, box)) ||
			pointsToSvgPath(points, mapPoint ? ['X', 'Y'] : entry.order);
		if (!path) continue;
		const brush = brushes.get(entry.brushRef) ?? DEFAULT_BRUSH;
		const pressures = pointsToPressures(
			entry.points,
			entry.order,
			Boolean(entry.authored),
			entry.force,
		);
		const tilt = pointsToTilt(entry.points, entry.order);
		const timestamps = entry.untimed ? [] : pointsToTimestamps(entry.points, entry.order);
		const { pointTimestamps, ...strokeTiming } = timing.read(entry.trace, timestamps);
		strokes.push({
			...brush,
			path,
			...(pressures.length > 0 ? { pressures } : {}),
			...(tilt && tilt.encoding === 'azimuthAltitude'
				? {
						tiltAngles: tilt.angles,
						tiltMagnitudes: tilt.magnitudes,
						tiltEncoding: 'azimuthAltitude',
					}
				: tilt
					? { tiltAngles: tilt.angles, tiltMagnitudes: tilt.magnitudes }
					: {}),
			...strokeTiming,
			...(pointTimestamps.length > 0 ? { pointTimestamps } : {}),
		});
	}
	return { strokes, rawXml: data, truncated };
}

/**
 * The size of the box this library's authored paths were drawn in (`pva:spaceWidth`,
 * `pva:spaceHeight` on the root), so a part read into a box of another size (the drawing was
 * resized after it was saved) is stretched onto it, as Office stretches the traces.
 */
function authoredSpace(root: unknown): { width: number; height: number } | undefined {
	if (!root || typeof root !== 'object') return undefined;
	const width = Number(nsAttr(root as InkXmlObject, 'spaceWidth'));
	const height = Number(nsAttr(root as InkXmlObject, 'spaceHeight'));
	return width > 0 && height > 0 ? { width, height } : undefined;
}

const AUTHORED_POINT =
	/([ML])\s*(-?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?)[,\s]+(-?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?)/giu;

/** An authored path stretched from the space it was drawn in onto `box`, else as it is. */
function scaleAuthoredPath(
	path: string,
	space: { width: number; height: number } | undefined,
	box: InkTargetBox | undefined,
): string {
	if (!space || !box || (space.width === box.width && space.height === box.height)) return path;
	const sx = box.width / space.width;
	const sy = box.height / space.height;
	const round = (value: number) => Math.round(value * 100) / 100;
	return path.replace(
		AUTHORED_POINT,
		(_, command: string, x: string, y: string) =>
			`${command} ${round(Number(x) * sx)} ${round(Number(y) * sy)}`,
	);
}

/** A channel's declared range (InkML `channel` `min`/`max`), when it declares a `max`. */
function channelRange(
	format: InkXmlObject | undefined,
	name: string,
): { min: number; max: number } | undefined {
	if (!format) return undefined;
	const channel = ensureArray(nsGet(format, 'channel')).find(
		(item) => String(nsAttr(item, 'name') ?? '').toUpperCase() === name,
	);
	if (!channel) return undefined;
	const max = Number(nsAttr(channel, 'max'));
	const declaredMin = Number(nsAttr(channel, 'min') ?? 0);
	const min = isFiniteNumber(declaredMin) ? declaredMin : 0;
	return isFiniteNumber(max) && max > min ? { min, max } : undefined;
}

/** Re-project decoded channel values, keeping every non-XY channel in place. */
function mapDecodedPoints(
	points: readonly (readonly number[])[],
	channelOrder: readonly string[],
	mapPoint: (x: number, y: number) => [number, number],
): number[][] {
	const xi = Math.max(channelOrder.indexOf('X'), 0);
	const yi = channelOrder.indexOf('Y') >= 0 ? channelOrder.indexOf('Y') : 1;
	// A point missing a channel maps through NaN, as the arithmetic on `undefined` did.
	return points.map((point) => mapPoint(point[xi] ?? NaN, point[yi] ?? NaN));
}

/**
 * Index every `<brush>` in the part by id, wherever it sits.
 *
 * PowerPoint nests its brushes inside `<inkml:definitions>`, so a direct-child
 * lookup found none of them and every real stroke fell back to a 1 px black
 * default. Brush measurements carry their own `units` attribute (PowerPoint
 * writes `units="cm"`), so a raw `Number(value)` would have produced a 0.05 px
 * stroke even once the brush was found.
 */
function collectBrushes(nodes: readonly InkXmlObject[]): Map<string, BrushStyle> {
	const brushes = new Map<string, BrushStyle>();
	for (const brush of nodes) {
		if (typeof brush === 'string') {
			continue;
		}
		const properties = new Map<string, { value: unknown; units: unknown }>();
		for (const property of ensureArray(nsGet(brush, 'brushProperty'))) {
			properties.set(String(nsAttr(property, 'name') ?? ''), {
				value: nsAttr(property, 'value'),
				units: nsAttr(property, 'units'),
			});
		}
		const size = properties.get('width') ?? properties.get('height');
		const width = size
			? inkLengthToPx(Number(size.value), size.units === undefined ? undefined : String(size.units))
			: Number.NaN;
		const id = String(nsAttr(brush, 'id') ?? '');
		const extras = [...properties]
			.filter(
				([name, { value }]) =>
					!['width', 'color', 'transparency', 'opacity'].includes(name) &&
					// A height equal to the width is the round tip the writer always writes.
					!(name === 'height' && size && value === size.value),
			)
			.map(([name, { value, units }]) => ({
				name,
				value: String(value ?? ''),
				...(units === undefined ? {} : { units: String(units) }),
			}));
		const highlighter = properties.get('rasterOp')?.value === 'maskPen';
		brushes.set(id, {
			...(highlighter ? { tool: 'highlighter' as const } : {}),
			...(extras.length ? { brushExtras: extras } : {}),
			color: String(properties.get('color')?.value ?? DEFAULT_BRUSH.color),
			// Rounded, so a width this library wrote in centimetres reads back as the same pixels.
			width:
				isFiniteNumber(width) && width > 0 ? Math.round(width * 1e6) / 1e6 : DEFAULT_BRUSH.width,
			opacity: brushOpacity(properties),
		});
	}
	return brushes;
}

/**
 * Stroke alpha. The library's own authored parts carry a direct `opacity`
 * (0..1); PowerPoint instead writes InkML's `transparency`, an integer where 0
 * is opaque and 255 is invisible (its highlighter pen is the usual producer).
 */
function brushOpacity(properties: Map<string, { value: unknown; units: unknown }>): number {
	const direct = Number(properties.get('opacity')?.value);
	if (isFiniteNumber(direct)) {
		return Math.min(1, Math.max(0, direct));
	}
	const transparency = Number(properties.get('transparency')?.value);
	if (isFiniteNumber(transparency)) {
		return 1 - Math.min(255, Math.max(0, transparency)) / 255;
	}
	return DEFAULT_BRUSH.opacity;
}
