import { isFiniteNumber } from '../../units/finite';
import type { InkStroke } from './types';

export function strokeStart(stroke: InkStroke): number | undefined {
	return isFiniteNumber(stroke.startTimestamp)
		? stroke.startTimestamp
		: stroke.pointTimestamps?.find(isFiniteNumber);
}

export function resolveInkClock(strokes: readonly InkStroke[]) {
	for (const stroke of strokes) {
		const start = strokeStart(stroke);
		if (start === undefined) continue;
		const date = new Date(Math.round(start));
		if (isFiniteNumber(date.getTime())) {
			return { base: date.getTime(), timeString: date.toISOString() };
		}
	}
	return undefined;
}

export function timestampOffset(sample: number, base: number): string {
	return String(Math.round(sample - base));
}
