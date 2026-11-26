import type { XmlObject } from '../types';
import type { InkMlIndex } from './inkml-limits';
import { ensureArray, nsAttr, nsGet } from './inkml-trace-decode';

function numberAttribute(node: XmlObject, key: string): number | undefined {
	const value = nsAttr(node, key);
	if (typeof value !== 'number' && (typeof value !== 'string' || !value.trim())) return undefined;
	const number = Number(value);
	return Number.isFinite(number) ? number : undefined;
}

function reference(node: XmlObject, key: string): string {
	const value = nsAttr(node, key);
	return typeof value === 'string' && value.startsWith('#') ? value.slice(1) : '';
}

/** Decode trace and context clocks from InkML 4.4 (https://www.w3.org/TR/InkML/#timestamps). */
export function inkMlTiming(index: InkMlIndex) {
	const timestamps = new Map<string, XmlObject>();
	const contexts = new Map<string, XmlObject>();
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
	const clockCache = new Map<XmlObject, number | undefined>();
	function clock(stamp: XmlObject | undefined, remaining = 16): number | undefined {
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
	const format = index.traceFormats[0];
	const channel =
		format &&
		ensureArray(nsGet(format, 'channel')).find(
			(value) => String(nsAttr(value, 'name')).toUpperCase() === 'T',
		);
	const unit = channel ? nsAttr(channel, 'units') : undefined;
	const channelBase = channel ? clock(timestamps.get(reference(channel, 'respectTo'))) : undefined;
	const scale = unit === 's' ? 1000 : unit === 'us' ? 0.001 : 1;
	return (trace: XmlObject | string, samples: number[]) => {
		if (typeof trace !== 'object') return { pointTimestamps: samples };
		const context = contexts.get(reference(trace, 'contextRef'));
		const contextStamp = context ? ensureArray(nsGet(context, 'timestamp'))[0] : undefined;
		const base = clock(timestamps.get(reference(trace, 'timestampRef'))) ?? clock(contextStamp);
		const offset = numberAttribute(trace, 'timeOffset');
		const startTimestamp =
			base === undefined && offset === undefined ? undefined : (base ?? 0) + (offset ?? 0);
		const durationMs = numberAttribute(trace, 'duration');
		return {
			pointTimestamps: samples.map(
				(sample) => sample * scale + (channelBase ?? startTimestamp ?? 0),
			),
			...(startTimestamp !== undefined ? { startTimestamp } : {}),
			...(durationMs !== undefined && durationMs > 0 ? { durationMs } : {}),
		};
	};
}
