import type { InkStroke } from './types';

/** Pressure that varies; this library writes a constant 0.5 for a stroke drawn without any. */
const pressureCount = (values: readonly number[] | undefined): number =>
	values?.some((value) => value !== 0.5) ? values.length : 0;

const sameNumbers = (a: readonly number[] | undefined, b: readonly number[] | undefined) =>
	(a ?? []).length === (b ?? []).length && (a ?? []).every((value, index) => value === b?.[index]);

/** Whether two stroke lists draw the same ink (paths, brushes, timing and per-point channels). */
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
				stroke.startTimestamp === other.startTimestamp &&
				stroke.durationMs === other.durationMs &&
				pressureCount(stroke.pressures) === pressureCount(other.pressures) &&
				(stroke.tiltAngles ?? []).length === (other.tiltAngles ?? []).length &&
				sameNumbers(stroke.pointTimestamps, other.pointTimestamps)
			);
		})
	);
}
