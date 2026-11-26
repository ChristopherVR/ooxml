import type { InkCoordinate } from './ink-geometry';

export interface InkReplayTiming {
	pointTimestamps?: readonly number[] | undefined;
	startTimestamp?: number | undefined;
	durationMs?: number | undefined;
}

export interface InkReplayEntry {
	strokeIndex: number;
	startOffsetMs: number;
	durationMs: number;
}

const DEFAULT_DURATION_MS = 600;
const DEFAULT_GAP_MS = 200;
const MAX_IDLE_GAP_MS = 3000;
const MAX_STROKE_DURATION_MS = 60_000;

/** Turn a fraction of the recorded time into a fraction of the drawn path length. */
export function inkReplayProgress(
	points: readonly InkCoordinate[],
	timestamps: readonly number[] | undefined,
): ((elapsed: number) => number) | undefined {
	const first = timestamps?.[0];
	const last = timestamps?.at(-1);
	if (
		!timestamps ||
		timestamps.length !== points.length ||
		first === undefined ||
		last === undefined
	) {
		return undefined;
	}
	const span = last - first;
	if (!Number.isFinite(span) || span <= 0) return undefined;
	const lengths = [0];
	let total = 0;
	let previousTime = first;
	for (let index = 1; index < points.length; index++) {
		const previous = points[index - 1];
		const point = points[index];
		const time = timestamps[index];
		if (
			!previous ||
			!point ||
			time === undefined ||
			!Number.isFinite(time) ||
			time < previousTime
		) {
			return undefined;
		}
		total += Math.hypot(point.x - previous.x, point.y - previous.y);
		lengths.push(total);
		previousTime = time;
	}
	if (!Number.isFinite(total) || total === 0) return undefined;
	return (elapsed) => {
		if (elapsed <= 0 || elapsed >= 1) return elapsed;
		const time = first + elapsed * span;
		let low = 0;
		let high = timestamps.length - 1;
		while (high - low > 1) {
			const middle = (low + high) >>> 1;
			if ((timestamps[middle] ?? last) <= time) low = middle;
			else high = middle;
		}
		const before = timestamps[low] ?? first;
		const after = timestamps[high] ?? last;
		const fraction = after > before ? (time - before) / (after - before) : 1;
		const start = lengths[low] ?? 0;
		const end = lengths[high] ?? total;
		return (start + fraction * (end - start)) / total;
	};
}

export function slideInkReplayTimeline(timings: readonly InkReplayTiming[]): InkReplayEntry[] {
	const steps = timings.map((timing, strokeIndex) => {
		const first = timing.pointTimestamps?.[0];
		const last = timing.pointTimestamps?.at(-1);
		const span = first !== undefined && last !== undefined ? last - first : undefined;
		const duration =
			span !== undefined && Number.isFinite(span) && span > 0 ? span : timing.durationMs;
		const time = Number.isFinite(first) ? first : timing.startTimestamp;
		return {
			strokeIndex,
			time: Number.isFinite(time) ? time : undefined,
			durationMs:
				duration !== undefined && Number.isFinite(duration)
					? Math.max(60, Math.min(MAX_STROKE_DURATION_MS, duration))
					: DEFAULT_DURATION_MS,
		};
	});
	if (steps.every(({ time }) => time !== undefined)) {
		steps.sort((a, b) => (a.time ?? 0) - (b.time ?? 0));
	}
	let offset = 0;
	let previousTime: number | undefined;
	let previousDuration = 0;
	return steps.map(({ strokeIndex, time, durationMs }) => {
		if (previousDuration > 0) {
			const gap =
				previousTime !== undefined && time !== undefined && time > previousTime
					? Math.max(0, Math.min(MAX_IDLE_GAP_MS, time - previousTime - previousDuration))
					: DEFAULT_GAP_MS;
			offset += previousDuration + gap;
		}
		previousTime = time;
		previousDuration = durationMs;
		return { strokeIndex, startOffsetMs: offset, durationMs };
	});
}
