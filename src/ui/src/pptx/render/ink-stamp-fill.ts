/**
 * Close the gaps between the per-point stamps of a pressure or tilt stroke.
 *
 * Pressure circles and nib marks are drawn one per sample point, so a stroke whose samples sit
 * further apart than its own radius (a fast pen stroke, or ink reloaded from a part written with
 * few points) painted as a row of separate dots instead of a line. Inserting interpolated stamps
 * between neighbours keeps the body connected; every binding paints the same descriptor, so the
 * fix reaches all of them.
 *
 * @module render/ink-stamp-fill
 */
import type { PressureCircle } from './ink-rendering';
import type { NibMark } from './ink-tilt-nib';

/** Stamps sit at most this fraction of the smaller neighbouring radius apart. */
const SPACING = 0.5;
/** A stroke never grows past this many stamps; beyond it the spacing widens instead. */
const MAX_STAMPS = 8192;
const MIN_STEP = 0.25;

/** Fill the gaps between neighbouring pressure circles. */
export function fillCircleGaps(circles: PressureCircle[]): PressureCircle[] {
	return fillGaps(
		circles,
		(c) => c.r,
		(a, b, t) => ({
			cx: a.cx + (b.cx - a.cx) * t,
			cy: a.cy + (b.cy - a.cy) * t,
			r: a.r + (b.r - a.r) * t,
		}),
	);
}

/** Fill the gaps between neighbouring nib marks, turning the nib the short way round. */
export function fillNibGaps(marks: NibMark[]): NibMark[] {
	return fillGaps(
		marks,
		(m) => Math.min(m.rTilt, m.rPerp),
		(a, b, t) => {
			const turn = b.rotationDeg - a.rotationDeg;
			const wrapped = turn - 360 * Math.round(turn / 360);
			return {
				cx: a.cx + (b.cx - a.cx) * t,
				cy: a.cy + (b.cy - a.cy) * t,
				rTilt: a.rTilt + (b.rTilt - a.rTilt) * t,
				rPerp: a.rPerp + (b.rPerp - a.rPerp) * t,
				rotationDeg: a.rotationDeg + wrapped * t,
			};
		},
	);
}

interface Stamp {
	cx: number;
	cy: number;
}

function fillGaps<T extends Stamp>(
	stamps: T[],
	radius: (stamp: T) => number,
	between: (a: T, b: T, t: number) => T,
): T[] {
	if (stamps.length < 2) return stamps;
	let total = 0;
	for (let i = 1; i < stamps.length; i++) total += distance(stamps[i - 1]!, stamps[i]!);
	// Spread the budget evenly when a long, thin stroke would need more stamps than allowed.
	const floor = Math.max(MIN_STEP, total / MAX_STAMPS);
	const out: T[] = [stamps[0]!];
	for (let i = 1; i < stamps.length; i++) {
		const a = stamps[i - 1]!;
		const b = stamps[i]!;
		const step = Math.max(floor, Math.min(radius(a), radius(b)) * SPACING);
		const count = Math.ceil(distance(a, b) / step);
		for (let k = 1; k < count; k++) out.push(between(a, b, k / count));
		out.push(b);
	}
	return out;
}

function distance(a: Stamp, b: Stamp): number {
	return Math.hypot(b.cx - a.cx, b.cy - a.cy);
}
