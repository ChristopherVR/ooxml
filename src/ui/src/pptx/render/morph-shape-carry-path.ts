/**
 * Path and stop helpers of `morph-shape-carry`: an outline's commands, their
 * box-normalised interpolation, and the keyframe stops that bake it.
 *
 * @module render/morph-shape-carry-path
 */
import type { PptxElement } from 'ooxml-core/pptx';

import { easingSegment, timeAtProgress } from './morph-easing';
import { parseSvgPath, serializeSvgPath } from './morph-svg-path';
import type { SvgPathCommand } from './morph-types';
import { MORPH_EASING_POINTS } from './morph-types';
import { getResolvedShapeClipPath } from './shape-geometry';
import { outlinePathData } from './stroke-outline';

/** Fewest stops baked into a per-stop track (clip or outline); never further apart than this. */
export const SHAPE_CARRY_STOPS = 12;
/** Most stops a track may take, however much the outline's share of the box changes. */
const MAX_CARRY_STOPS = 32;
/**
 * Largest ratio the outline's share of the box may change by between two
 * stops. The re-pathed `d` is that share's reciprocal, which CSS interpolates
 * linearly from stop to stop; at 1.1 that is within 0.3% of the box.
 */
const MAX_STOP_RATIO = 1.1;

type Axis = 'x' | 'y' | 'keep';

/** Which axis an SVG path command's `index`th value lives on. */
function axisOf(type: string, index: number): Axis {
	const upper = type.toUpperCase();
	if (upper === 'H') return 'x';
	if (upper === 'V') return 'y';
	if (upper === 'A') {
		return (['x', 'y', 'keep', 'keep', 'keep', 'x', 'y'] as const)[index % 7] ?? 'keep';
	}
	return index % 2 === 0 ? 'x' : 'y';
}

export function outlineCommands(element: PptxElement): SvgPathCommand[] | undefined {
	const d = outlinePathData(getResolvedShapeClipPath(element), element.width, element.height);
	const commands = d ? parseSvgPath(d) : [];
	return commands.length > 0 ? commands : undefined;
}

/** Whether two outlines are the same commands with the same flags, so their vertices pair up. */
export function sameStructure(a: SvgPathCommand[], b: SvgPathCommand[]): boolean {
	return (
		a.length === b.length &&
		a.every((command, index) => {
			const other = b[index];
			return (
				other !== undefined &&
				other.type === command.type &&
				other.values.length === command.values.length &&
				command.values.every(
					(value, k) =>
						axisOf(command.type, k) !== 'keep' || Math.abs(value - (other.values[k] ?? 0)) < 1e-6,
				)
			);
		})
	);
}

export function lerp(a: number, b: number, p: number): number {
	return a + (b - a) * p;
}

/** The pair's geometry at eased progress `p`, scaled from box-normalised units by `kx`, `ky`. */
export function geometryAt(
	a: SvgPathCommand[],
	b: SvgPathCommand[],
	from: PptxElement,
	to: PptxElement,
	p: number,
	kx: number,
	ky: number,
): string {
	const commands = a.map((command, index) => ({
		type: command.type,
		values: command.values.map((value, k) => {
			const other = b[index]?.values[k] ?? value;
			const axis = axisOf(command.type, k);
			if (axis === 'keep') return value;
			return axis === 'x'
				? lerp(value / from.width, other / to.width, p) * kx
				: lerp(value / from.height, other / to.height, p) * ky;
		}),
	}));
	return serializeSvgPath(commands);
}

/** Scale every coordinate of a path by `kx`, `ky` (flags and rotations kept). */
export function scalePath(d: string, kx: number, ky: number): string {
	return serializeSvgPath(
		parseSvgPath(d).map((command) => ({
			type: command.type,
			values: command.values.map((value, k) => {
				const axis = axisOf(command.type, k);
				return axis === 'x' ? value * kx : axis === 'y' ? value * ky : value;
			}),
		})),
	);
}

export function round(value: number): number {
	return Math.round(value * 1000) / 1000;
}

/**
 * The eased progress of each stop: evenly spaced, plus, where the outline's
 * share of the box changes by a large ratio `u0 -> u1`, spaced geometrically in
 * that share so no two neighbours differ by more than {@link MAX_STOP_RATIO}.
 */
export function stopProgress(u0: number, u1: number): number[] {
	const stops = new Set<number>();
	for (let i = 0; i <= SHAPE_CARRY_STOPS; i++) stops.add(i / SHAPE_CARRY_STOPS);
	const ratio = u0 > 0 && u1 > 0 ? Math.max(u0 / u1, u1 / u0) : 1;
	const n = Math.min(
		MAX_CARRY_STOPS - SHAPE_CARRY_STOPS,
		Math.ceil(Math.log(ratio) / Math.log(MAX_STOP_RATIO)),
	);
	for (let i = 1; i < n; i++) {
		stops.add(round((u0 * Math.pow(u1 / u0, i / n) - u0) / (u1 - u0)));
	}
	const sorted = [...stops].sort((a, b) => a - b);
	return sorted.filter((p, i) => i === 0 || p - (sorted[i - 1] ?? 0) > 0.002 || p === 1);
}

/** One keyframe per stop, each carrying its slice of the morph curve to the next. */
export function bakeStops(
	progress: readonly number[],
	declarationsAt: (p: number) => string[],
): string {
	const lines: string[] = [];
	progress.forEach((p, i) => {
		const next = progress[i + 1];
		const segment = next === undefined ? undefined : easingSegment(MORPH_EASING_POINTS, p, next);
		const time = p <= 0 ? 0 : p >= 1 ? 100 : round(timeAtProgress(p, MORPH_EASING_POINTS) * 100);
		const body = declarationsAt(p).map((declaration) => `\t\t${declaration};`);
		if (segment) body.push(`\t\tanimation-timing-function: ${segment.easing};`);
		lines.push(`\t${time}% {\n${body.join('\n')}\n\t}`);
	});
	return lines.join('\n');
}
