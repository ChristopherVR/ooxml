import type { InkXmlObject } from './types';
import { countInkTraceSamples, type InkMlIndex } from './inkml-limits';
import { ensureArray, nsAttr, nsGet } from './inkml-trace-decode';

function numberAttribute(node: InkXmlObject, key: string): number | undefined {
	const value = nsAttr(node, key);
	if (typeof value !== 'number' && (typeof value !== 'string' || !value.trim())) return undefined;
	const number = Number(value);
	return Number.isFinite(number) ? number : undefined;
}

function reference(node: InkXmlObject, key: string): string {
	const value = nsAttr(node, key);
	return typeof value === 'string' && value.startsWith('#') ? value.slice(1) : '';
}

/** Decode trace and context clocks from InkML 4.4 (https://www.w3.org/TR/InkML/#timestamps). */
export function inkMlTiming(index: InkMlIndex) {
	const timestamps = new Map<string, InkXmlObject>();
	const contexts = new Map<string, InkXmlObject>();
	for (const stamp of index.timestamps) {
		if (typeof stamp !== 'object') continue;
		const id = nsAttr(stamp, 'id');
		if (typeof id === 'string') timestamps.set(id, stamp);
	}
	for (const context of index.contexts) {
		if (typeof context !== 'object') continue;
		const id = nsAttr(context, 'id');
		if (typeof id === 'string') contexts.set(id, context);
	}
	const clockCache = new Map<InkXmlObject, number | undefined>();
	function clock(stamp: InkXmlObject | undefined, remaining = 16): number | undefined {
		if (!stamp || remaining === 0) return undefined;
		if (clockCache.has(stamp)) return clockCache.get(stamp);
		const text = nsAttr(stamp, 'timeString');
		const date =
			typeof text === 'string' && text.length < 64
				? Date.parse(/(?:Z|[+-]\d\d:\d\d)$/u.test(text) ? text : `${text}Z`)
				: NaN;
		const time =
			numberAttribute(stamp, 'time') ??
			(Number.isFinite(date)
				? date
				: clock(timestamps.get(reference(stamp, 'timestampRef')), remaining - 1));
		const offset = numberAttribute(stamp, 'timeOffset') ?? 0;
		const value = time === undefined ? undefined : time + offset;
		clockCache.set(stamp, value);
		return value;
	}
	function contextFor(trace: InkXmlObject | string): InkXmlObject | undefined {
		return typeof trace === 'object' ? contexts.get(reference(trace, 'contextRef')) : undefined;
	}
	function formatFor(trace: InkXmlObject | string): InkXmlObject | undefined {
		const context = contextFor(trace);
		const source = context && ensureArray(nsGet(context, 'inkSource'))[0];
		return (
			(context && ensureArray(nsGet(context, 'traceFormat'))[0]) ??
			(source && ensureArray(nsGet(source, 'traceFormat'))[0]) ??
			index.traceFormats[0]
		);
	}
	const formatClocks = new Map<InkXmlObject, { channelBase: number | undefined; scale: number }>();
	function channelClock(format: InkXmlObject | undefined) {
		if (!format) return { channelBase: undefined, scale: 1 };
		const cached = formatClocks.get(format);
		if (cached) return cached;
		const channel = ensureArray(nsGet(format, 'channel')).find(
			(value) => String(nsAttr(value, 'name')).toUpperCase() === 'T',
		);
		const unit = channel ? nsAttr(channel, 'units') : undefined;
		const result = {
			channelBase: channel ? clock(timestamps.get(reference(channel, 'respectTo'))) : undefined,
			scale: unit === 's' ? 1000 : unit === 'us' ? 0.001 : 1,
		};
		formatClocks.set(format, result);
		return result;
	}
	function read(trace: InkXmlObject | string, samples: number[]) {
		const { channelBase, scale } = channelClock(formatFor(trace));
		if (typeof trace !== 'object') {
			return { pointTimestamps: samples.map((sample) => sample * scale + (channelBase ?? 0)) };
		}
		const context = contextFor(trace);
		const contextStamp = context ? ensureArray(nsGet(context, 'timestamp'))[0] : undefined;
		const base = clock(timestamps.get(reference(trace, 'timestampRef'))) ?? clock(contextStamp);
		const offset = numberAttribute(trace, 'timeOffset');
		const startTimestamp =
			base === undefined && offset === undefined ? undefined : (base ?? 0) + (offset ?? 0);
		const durationMs = numberAttribute(trace, 'duration');
		const extension = trace['@_pva:sampleTimes'];
		let pointTimestamps = samples.map(
			(sample) => sample * scale + (channelBase ?? startTimestamp ?? 0),
		);
		if (
			typeof extension === 'string' &&
			base !== undefined &&
			countInkTraceSamples(extension) !== undefined
		) {
			const offsets = extension.split(',').map(Number);
			if (offsets.every(Number.isSafeInteger)) {
				pointTimestamps = offsets.map((offset) => offset + base);
			}
		}
		return {
			pointTimestamps,
			...(startTimestamp !== undefined ? { startTimestamp } : {}),
			...(durationMs !== undefined && durationMs > 0 ? { durationMs } : {}),
		};
	}
	return { formatFor, read };
}
