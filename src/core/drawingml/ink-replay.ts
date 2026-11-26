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
