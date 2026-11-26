import { MAX_INK_SAMPLES } from '../../../drawingml/ink-geometry';
import type { XmlObject } from '../types';
import { ensureArray, localNameOf } from './inkml-trace-decode';

export const MAX_INK_PART_CHARS = 8 * 1024 * 1024;
export const MAX_INK_TRACES = 1000;
export const MAX_INK_PART_SAMPLES = 100_000;
export const MAX_INK_TRACE_CHARS = MAX_INK_SAMPLES * 64;
const MAX_INK_XML_DEPTH = 64;
const MAX_INK_CHANNELS = 32;

export interface InkMlIndex {
	traces: XmlObject[];
	brushes: XmlObject[];
	contexts: XmlObject[];
	timestamps: XmlObject[];
	traceFormats: XmlObject[];
}

/** Index the XML once, including a budget for callers that already hold a parsed tree. */
export function indexInkMl(data: XmlObject): InkMlIndex | undefined {
	const index: InkMlIndex = {
		traces: [],
		brushes: [],
		contexts: [],
		timestamps: [],
		traceFormats: [],
	};
	let remaining = MAX_INK_PART_CHARS;
	function visit(node: unknown, depth: number): boolean {
		if (depth > MAX_INK_XML_DEPTH) return false;
		if (typeof node !== 'object' || node === null) {
			remaining -= typeof node === 'string' ? node.length : 1;
			return remaining >= 0;
		}
		if (Array.isArray(node)) {
			for (const item of node) {
				if (--remaining < 0 || !visit(item, depth + 1)) return false;
			}
			return true;
		}
		const object = node as XmlObject;
		for (const key in object) {
			if (!Object.hasOwn(object, key)) continue;
			remaining -= key.length + 1;
			if (remaining < 0) return false;
			const value = object[key];
			if (!key.startsWith('@_') && key !== '#text') {
				const entries = ensureArray(value);
				switch (localNameOf(key)) {
					case 'trace':
						if (index.traces.length + entries.length > MAX_INK_TRACES) return false;
						index.traces.push(...entries);
						break;
					case 'brush':
					case 'context':
					case 'timestamp':
					case 'traceFormat': {
						const name = localNameOf(key);
						const target =
							name === 'brush'
								? index.brushes
								: name === 'context'
									? index.contexts
									: name === 'timestamp'
										? index.timestamps
										: index.traceFormats;
						for (const entry of entries) {
							if (entry && typeof entry === 'object') target.push(entry);
						}
						break;
					}
					case 'channel':
						if (entries.length > MAX_INK_CHANNELS) return false;
						break;
				}
			}
			if (!visit(value, depth + 1)) return false;
		}
		return true;
	}
	return visit(data, 0) ? index : undefined;
}

/** Inspect trace text before allocating token or point arrays. */
export function countInkTraceSamples(text: string): number | undefined {
	if (text.length > MAX_INK_TRACE_CHARS) return undefined;
	let points = 1;
	for (let offset = text.indexOf(','); offset >= 0; offset = text.indexOf(',', offset + 1)) {
		if (++points > MAX_INK_SAMPLES) return undefined;
	}
	return points;
}

/** Authored SVG can contain points even when the trace text is empty. */
export function countAuthoredInkPathSamples(path: string): number | undefined {
	if (!path) return 0;
	if (path.length > MAX_INK_TRACE_CHARS) return undefined;
	let commands = 0;
	let numbers = 0;
	let singleCoordinate = false;
	const tokens = /[a-df-z]|[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?/giu;
	let match = tokens.exec(path);
	while (match) {
		if (/^[a-z]/iu.test(match[0])) {
			commands++;
			singleCoordinate = /^[hv]$/iu.test(match[0]);
		} else {
			numbers += singleCoordinate ? 2 : 1;
		}
		if (commands > MAX_INK_SAMPLES || numbers > MAX_INK_SAMPLES * 2) return undefined;
		match = tokens.exec(path);
	}
	return Math.max(commands, Math.ceil(numbers / 2));
}
