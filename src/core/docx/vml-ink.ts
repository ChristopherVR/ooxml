// Legacy Word ink: a VML `v:shape` holding an `o:ink` element (Word 2003 and earlier, and files
// Word converts from them). The strokes themselves are Ink Serialized Format in `o:ink/@i`, which
// this package does not decode; the shape's own `path` (ECMA-376 Part 4, VML `v:shape/@path`)
// holds the same strokes as lines and curves in the shape's `coordsize` space, so that is what is
// drawn. The ink is shown read only: the Draw tab leaves it alone and the VML is saved as read.
import { parseVmlColor, parseVmlOpacity } from '../color/vml-color';
import type { InkStroke } from '../drawingml/inkml/types';
import { first } from '../xml/index';
import type { XmlElement } from './xml';

const OFFICE_NS = 'urn:schemas-microsoft-com:office:office';
const VML_NS = 'urn:schemas-microsoft-com:vml';

/** Legacy VML ink: its strokes in a `width` x `height` pixel space (the shape's box). */
export interface LegacyVmlInk {
	strokes: InkStroke[];
	width: number;
	height: number;
}

/** A drawing holds at most this many strokes and points; more is left out of the drawing. */
const MAX_STROKES = 2048;
const MAX_POINTS = 65_536;
/** Points a VML curve segment is drawn with. */
const CURVE_STEPS = 8;
const PX_PER_UNIT: Record<string, number> = {
	px: 1,
	pt: 96 / 72,
	in: 96,
	cm: 96 / 2.54,
	mm: 96 / 25.4,
	emu: 96 / 914_400,
};

/** Whether a VML shape is ink (`o:ink` among its children). */
export const isVmlInkShape = (shape: XmlElement | undefined): shape is XmlElement =>
	Boolean(shape && first(shape, 'ink', OFFICE_NS));

function pair(value: string | null, fallback: [number, number]): [number, number] {
	const [a, b] = (value ?? '').split(',').map((part) => Number(part.trim() || 0));
	return Number.isFinite(a) && Number.isFinite(b) && value ? [a!, b!] : fallback;
}

function lengthPx(value: string | null | undefined, fallback: number): number {
	const match = /^\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+))\s*(px|pt|in|cm|mm|emu)?\s*$/i.exec(value ?? '');
	if (!match) return fallback;
	// A weight without a unit is read as points, the unit of VML's default weight (`1pt`).
	const px = Number(match[1]) * (PX_PER_UNIT[(match[2] ?? 'pt').toLowerCase()] ?? 1);
	return Number.isFinite(px) && px > 0 ? px : fallback;
}

type Point = { x: number; y: number };

/**
 * The sub-paths of a VML path (`m`, `l`, `c`, `t`, `r`, `v`, `x`, `e`; other commands end the
 * sub-path they are in) as polylines, in the path's own coordinates. Curves are sampled.
 */
export function vmlPathPolylines(path: string): Point[][] {
	const tokens = path.match(/[a-z]+|[-+]?\d*\.?\d+|,/gi) ?? [];
	const lines: Point[][] = [];
	let current: Point[] | undefined;
	let at: Point = { x: 0, y: 0 };
	let points = 0;
	let index = 0;
	// Commas with nothing between them stand for 0, as VML's grammar has it.
	const numbers = (): number[] => {
		const out: number[] = [];
		let pending = false;
		while (index < tokens.length && !/^[a-z]/i.test(tokens[index]!)) {
			const token = tokens[index++]!;
			if (token === ',') {
				if (!pending) out.push(0);
				pending = false;
			} else {
				out.push(Number(token));
				pending = true;
			}
		}
		return out.filter(Number.isFinite);
	};
	const push = (point: Point): void => {
		if (points >= MAX_POINTS) return;
		if (!current) {
			current = [at];
			lines.push(current);
		}
		current.push(point);
		at = point;
		points++;
	};
	while (index < tokens.length && lines.length <= MAX_STROKES) {
		const command = tokens[index++]!.toLowerCase();
		if (!/^[a-z]/.test(command)) continue;
		const args = numbers();
		if (command === 'm' || command === 't') {
			for (let k = 0; k + 1 < args.length; k += 2) {
				const next =
					command === 'm'
						? { x: args[k]!, y: args[k + 1]! }
						: { x: at.x + args[k]!, y: at.y + args[k + 1]! };
				if (k === 0) {
					current = undefined;
					at = next;
				} else push(next);
			}
		} else if (command === 'l' || command === 'r') {
			for (let k = 0; k + 1 < args.length; k += 2)
				push(
					command === 'l'
						? { x: args[k]!, y: args[k + 1]! }
						: { x: at.x + args[k]!, y: at.y + args[k + 1]! },
				);
		} else if (command === 'c' || command === 'v') {
			for (let k = 0; k + 5 < args.length; k += 6) {
				const o = command === 'v' ? at : { x: 0, y: 0 };
				const start = at;
				const [c1, c2, end] = [0, 2, 4].map((j) => ({
					x: o.x + args[k + j]!,
					y: o.y + args[k + j + 1]!,
				})) as [Point, Point, Point];
				for (let step = 1; step <= CURVE_STEPS; step++) {
					const t = step / CURVE_STEPS;
					const u = 1 - t;
					push({
						x:
							u * u * u * start.x + 3 * u * u * t * c1.x + 3 * u * t * t * c2.x + t * t * t * end.x,
						y:
							u * u * u * start.y + 3 * u * u * t * c1.y + 3 * u * t * t * c2.y + t * t * t * end.y,
					});
				}
			}
		} else {
			// `x` closes, `e` ends; arcs and the rest are not ink commands: the sub-path ends there.
			if (command === 'x' && current?.[0]) push(current[0]);
			current = undefined;
		}
	}
	return lines.filter((line) => line.length > 1).slice(0, MAX_STROKES);
}

const round = (value: number): number => Math.round(value * 100) / 100;

/**
 * The strokes of a legacy VML ink shape drawn in a `width` x `height` box, or undefined when the
 * shape is not ink or its path holds no lines.
 */
export function parseVmlInk(
	shape: XmlElement | undefined,
	width: number,
	height: number,
): LegacyVmlInk | undefined {
	if (!isVmlInkShape(shape) || !(width > 0 && height > 0)) return undefined;
	const lines = vmlPathPolylines(shape.getAttribute('path') ?? '');
	if (!lines.length) return undefined;
	const [cw, ch] = pair(shape.getAttribute('coordsize'), [1000, 1000]);
	const [ox, oy] = pair(shape.getAttribute('coordorigin'), [0, 0]);
	const kx = cw ? width / cw : 1;
	const ky = ch ? height / ch : 1;
	const stroke = first(shape, 'stroke', VML_NS);
	// A colour may name its system colour index after it (`red [10]`).
	const named = stroke?.getAttribute('color') ?? shape.getAttribute('strokecolor') ?? 'black';
	const color = parseVmlColor(named.replace(/\s*\[.*$/, ''))?.toUpperCase() ?? '#000000';
	const opacity = parseVmlOpacity(stroke?.getAttribute('opacity') ?? '') ?? 1;
	const weight = lengthPx(
		stroke?.getAttribute('weight') ?? shape.getAttribute('strokeweight'),
		96 / 72,
	);
	return {
		strokes: lines.map((line) => ({
			path: line
				.map((p, n) => `${n ? 'L' : 'M'} ${round((p.x - ox) * kx)} ${round((p.y - oy) * ky)}`)
				.join(' '),
			color: /^#[\dA-F]{6}$/.test(color) ? color : '#000000',
			width: round(weight),
			opacity,
		})),
		width,
		height,
	};
}
