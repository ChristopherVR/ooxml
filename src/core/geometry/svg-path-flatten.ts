/**
 * Flatten an SVG path `d` string into closed polygon loops (2D points), for
 * consumers that need a concrete outline rather than curve primitives (the
 * three.js SmartArt renderer extrudes/fills a shape's outline as a flat
 * polygon; sampling curves here keeps that consumer three.js-agnostic and
 * pure/testable).
 *
 * Supports M/L/H/V/C/Q/S/A/Z, both absolute and relative variants, and
 * multiple subpaths (each additional `M` starts a new loop; the first loop is
 * the outer boundary, subsequent loops are holes - the same convention as
 * `THREE.Shape`/`THREE.Path`.holes).
 *
 * @module render/svg-path-flatten
 */

import type { Point2 } from './svg-path-types.js';
export type { Point2 } from './svg-path-types.js';
import { sampleCubic, sampleQuadratic, sampleArc } from './svg-path-curves.js';
import { at } from './indexed.js';

const COMMAND_RE = /[MLCQZAHVSmlcqzahvs][^MLCQZAHVSmlcqzahvs]*/gu;
const NUMBER_RE = /-?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/gu;

interface Token {
	type: string;
	values: number[];
}

function tokenize(d: string): Token[] {
	const tokens: Token[] = [];
	const matches = d.match(COMMAND_RE);
	if (!matches) {
		return tokens;
	}
	for (const raw of matches) {
		const type = raw.charAt(0);
		const nums = raw.slice(1).match(NUMBER_RE);
		tokens.push({ type, values: nums ? nums.map((n) => Number.parseFloat(n)) : [] });
	}
	return tokens;
}

/**
 * Flatten an SVG path `d` string into one or more closed polygon loops.
 *
 * @param d - The path `d` attribute string (absolute or relative commands).
 * @param curveSegments - Sample count per full curve (arcs scale down for
 *   short sweeps). Default 16, generous enough for shapes rendered at
 *   on-screen SmartArt sizes.
 */
export function flattenSvgPath(d: string, curveSegments = 16): Point2[][] {
	if (!Number.isFinite(curveSegments) || curveSegments < 1) curveSegments = 16;
	curveSegments = Math.min(4096, Math.floor(curveSegments));
	const tokens = tokenize(d);
	const loops: Point2[][] = [];
	let current: Point2[] = [];
	let cursor: Point2 = { x: 0, y: 0 };
	let subpathStart: Point2 = { x: 0, y: 0 };
	let lastControl: Point2 | undefined;
	let lastCommand = '';

	const closeLoop = (): void => {
		if (current.length > 0) {
			loops.push(current);
		}
		current = [];
	};

	for (const token of tokens) {
		const type = token.type;
		const upper = type.toUpperCase();
		const relative = type !== upper;
		const v = token.values;
		// Incomplete commands cannot supply a point. Ignore them without leaking NaN coordinates.
		if (v.some((value) => !Number.isFinite(value)) || (upper === 'M' && v.length < 2)) continue;

		if (upper === 'M') {
			closeLoop();
			const x = relative ? cursor.x + at(v, 0) : at(v, 0);
			const y = relative ? cursor.y + at(v, 1) : at(v, 1);
			cursor = { x, y };
			subpathStart = cursor;
			current.push(cursor);
			// Extra coordinate pairs after the first M behave as implicit L.
			for (let i = 2; i + 1 < v.length; i += 2) {
				const lx = relative ? cursor.x + at(v, i) : at(v, i);
				const ly = relative ? cursor.y + at(v, i + 1) : at(v, i + 1);
				cursor = { x: lx, y: ly };
				current.push(cursor);
			}
		} else if (upper === 'L') {
			for (let i = 0; i + 1 < v.length; i += 2) {
				const x = relative ? cursor.x + at(v, i) : at(v, i);
				const y = relative ? cursor.y + at(v, i + 1) : at(v, i + 1);
				cursor = { x, y };
				current.push(cursor);
			}
		} else if (upper === 'H') {
			for (const value of v) {
				const x = relative ? cursor.x + value : value;
				cursor = { x, y: cursor.y };
				current.push(cursor);
			}
		} else if (upper === 'V') {
			for (const value of v) {
				const y = relative ? cursor.y + value : value;
				cursor = { x: cursor.x, y };
				current.push(cursor);
			}
		} else if (upper === 'C') {
			for (let i = 0; i + 5 < v.length; i += 6) {
				const p1 = relative
					? { x: cursor.x + at(v, i), y: cursor.y + at(v, i + 1) }
					: { x: at(v, i), y: at(v, i + 1) };
				const p2 = relative
					? { x: cursor.x + at(v, i + 2), y: cursor.y + at(v, i + 3) }
					: { x: at(v, i + 2), y: at(v, i + 3) };
				const p3 = relative
					? { x: cursor.x + at(v, i + 4), y: cursor.y + at(v, i + 5) }
					: { x: at(v, i + 4), y: at(v, i + 5) };
				sampleCubic(cursor, p1, p2, p3, curveSegments, current);
				cursor = p3;
				lastControl = p2;
			}
		} else if (upper === 'S') {
			for (let i = 0; i + 3 < v.length; i += 4) {
				const p1 =
					lastCommand === 'C' || lastCommand === 'S'
						? {
								x: 2 * cursor.x - (lastControl?.x ?? cursor.x),
								y: 2 * cursor.y - (lastControl?.y ?? cursor.y),
							}
						: cursor;
				const p2 = relative
					? { x: cursor.x + at(v, i), y: cursor.y + at(v, i + 1) }
					: { x: at(v, i), y: at(v, i + 1) };
				const p3 = relative
					? { x: cursor.x + at(v, i + 2), y: cursor.y + at(v, i + 3) }
					: { x: at(v, i + 2), y: at(v, i + 3) };
				sampleCubic(cursor, p1, p2, p3, curveSegments, current);
				cursor = p3;
				lastControl = p2;
			}
		} else if (upper === 'Q') {
			for (let i = 0; i + 3 < v.length; i += 4) {
				const p1 = relative
					? { x: cursor.x + at(v, i), y: cursor.y + at(v, i + 1) }
					: { x: at(v, i), y: at(v, i + 1) };
				const p2 = relative
					? { x: cursor.x + at(v, i + 2), y: cursor.y + at(v, i + 3) }
					: { x: at(v, i + 2), y: at(v, i + 3) };
				sampleQuadratic(cursor, p1, p2, curveSegments, current);
				cursor = p2;
				lastControl = p1;
			}
		} else if (upper === 'A') {
			for (let i = 0; i + 6 < v.length; i += 7) {
				const rx = at(v, i);
				const ry = at(v, i + 1);
				const rot = at(v, i + 2);
				const largeArc = at(v, i + 3) !== 0;
				const sweep = at(v, i + 4) !== 0;
				const end = relative
					? { x: cursor.x + at(v, i + 5), y: cursor.y + at(v, i + 6) }
					: { x: at(v, i + 5), y: at(v, i + 6) };
				sampleArc(cursor, rx, ry, rot, largeArc, sweep, end, curveSegments, current);
				cursor = end;
			}
		} else if (upper === 'Z') {
			cursor = subpathStart;
			current.push(cursor);
		}
		lastCommand = upper;
	}
	closeLoop();
	return loops;
}
