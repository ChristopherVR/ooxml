import type { InkStroke } from './types';

/** Pressure that varies; this library writes a constant 0.5 for a stroke drawn without any. */
const pressuresOf = (values: readonly number[] | undefined): readonly number[] =>
	values?.some((value) => value !== 0.5) ? values : [];

const sameNumbers = (a: readonly number[] | undefined, b: readonly number[] | undefined) =>
	(a ?? []).length === (b ?? []).length && (a ?? []).every((value, index) => value === b?.[index]);

const isHighlighter = (stroke: InkStroke) => stroke.tool === 'highlighter';

/** Brush extras as written: a highlighter's own tip and raster operation go without saying. */
function extrasOf(stroke: InkStroke): string {
	const extras = (stroke.brushExtras ?? []).filter(
		(extra) =>
			!(
				isHighlighter(stroke) &&
				((extra.name === 'tip' && extra.value === 'rectangle') ||
					(extra.name === 'rasterOp' && extra.value === 'maskPen'))
			),
	);
	return JSON.stringify(
		extras
			.map((extra) => [extra.name, extra.value, extra.units ?? ''])
			.sort((a, b) => String(a[0]).localeCompare(String(b[0]))),
	);
}

/**
 * Whether two stroke lists draw the same ink: paths, brushes (with the properties this model
 * does not hold), the pen (a highlighter or not, which is all a part records), timing, and the
 * pressure and tilt of every point.
 */
export function inkStrokesEqual(left: readonly InkStroke[], right: readonly InkStroke[]): boolean {
	return (
		left.length === right.length &&
		left.every((stroke, index) => {
			const other = right[index];
			return (
				other !== undefined &&
				stroke.path === other.path &&
				stroke.color === other.color &&
				stroke.width === other.width &&
				stroke.opacity === other.opacity &&
				isHighlighter(stroke) === isHighlighter(other) &&
				stroke.startTimestamp === other.startTimestamp &&
				stroke.durationMs === other.durationMs &&
				sameNumbers(pressuresOf(stroke.pressures), pressuresOf(other.pressures)) &&
				sameNumbers(stroke.tiltAngles, other.tiltAngles) &&
				sameNumbers(stroke.tiltMagnitudes, other.tiltMagnitudes) &&
				sameNumbers(stroke.pointTimestamps, other.pointTimestamps) &&
				extrasOf(stroke) === extrasOf(other)
			);
		})
	);
}
