import { at } from '../../geometry/indexed';
import { isFiniteNumber } from '../../units/finite';
import type { ChannelOrder } from './inkml-trace-decode';

/**
 * Decode a raw InkML trace string into per-point channel values.
 *
 * Points are comma-separated; the channel values within a point map
 * positionally to `channelOrder`. InkML permits value prefixes that switch a
 * channel into a difference encoding: `!` explicit (absolute), `'` single
 * difference (delta from the previous point), `"` second difference (delta of
 * the delta). The mode is sticky per channel until another prefix appears.
 *
 * The values are NOT reliably whitespace-separated, which is the trap here.
 * PowerPoint emits its traces in the compact form, where the sign doubles as
 * the separator and the prefix binds tight to its number. This is verbatim
 * PowerPoint SaveAs output for a three-point stroke:
 *
 *     100 200,'40'46,"0"-5,0-10
 *
 * A whitespace split reads `'40'46` and `0-10` as ONE token each and drops
 * every point after the first, which is exactly what used to happen: a real
 * inked slide decoded to a single point and rendered nothing. Values are
 * therefore scanned, not split.
 */
export function decodeTracePoints(text: string, channelOrder: ChannelOrder): number[][] {
	const points: number[][] = [];
	const modes: DiffMode[] = channelOrder.map(() => 'explicit');
	const lastValue: number[] = channelOrder.map(() => 0);
	const lastVelocity: number[] = channelOrder.map(() => 0);
	for (const rawPoint of text.split(',')) {
		const tokens = scanValueTokens(rawPoint, channelOrder.length);
		if (tokens.length === 0) {
			continue;
		}
		const decoded: number[] = [];
		for (let i = 0; i < tokens.length && i < channelOrder.length; i++) {
			// `lastValue` and `modes` have one entry per channel and `i` is below both lengths.
			const parsed = parseValueToken(at(tokens, i), at(modes, i));
			if (parsed === undefined) {
				decoded.push(at(lastValue, i));
				continue;
			}
			modes[i] = parsed.mode;
			const value = applyDiffMode(parsed, i, lastValue, lastVelocity);
			decoded.push(value);
		}
		if (decoded.length >= 2) {
			points.push(decoded);
		}
	}
	return points;
}

type DiffMode = 'explicit' | 'single' | 'double';

interface ParsedToken {
	value: number;
	mode: DiffMode;
	/** `*`: repeat the last value (explicit) or the last difference. */
	repeat?: boolean;
}

/**
 * One optional InkML encoding prefix followed by one signed decimal number.
 * Anchored on the number so that `0-10` yields two values and `'40'46` keeps
 * each prefix attached to the value it qualifies.
 */
const VALUE_TOKEN =
	/(?:(['"!])\s*)?(\?|\*|-?#[\da-fA-F]+|-?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)/gu;

/** Scan one comma-delimited point into its channel value tokens. */
function scanValueTokens(rawPoint: string, channelCount: number): string[] {
	const tokens: string[] = [];
	VALUE_TOKEN.lastIndex = 0;
	let match: RegExpExecArray | null;
	while (tokens.length < channelCount && (match = VALUE_TOKEN.exec(rawPoint)) !== null) {
		tokens.push(`${match[1] ?? ''}${match[2]}`);
	}
	return tokens;
}

/** Parse one whitespace-delimited channel token, honouring InkML mode prefixes. */
function parseValueToken(token: string, currentMode: DiffMode): ParsedToken | undefined {
	let mode = currentMode;
	let body = token;
	const prefix = token[0];
	if (prefix === '!') {
		mode = 'explicit';
		body = token.slice(1);
	} else if (prefix === "'") {
		mode = 'single';
		body = token.slice(1);
	} else if (prefix === '"') {
		mode = 'double';
		body = token.slice(1);
	}
	// `?`: the value is unknown, so the channel keeps its last value.
	if (body.length === 0 || body === '?') {
		return undefined;
	}
	// `*`: the same value (explicit) or the same difference as the last point.
	if (body === '*') {
		return { value: Number.NaN, mode, repeat: true };
	}
	// `#`: a hexadecimal integer.
	const hex = /^(-?)#([\da-fA-F]+)$/u.exec(body);
	const value = hex ? (hex[1] ? -1 : 1) * Number.parseInt(hex[2] ?? '', 16) : Number(body);
	return isFiniteNumber(value) ? { value, mode } : undefined;
}

/** Fold a parsed token into an absolute value using its (sticky) difference mode. */
function applyDiffMode(
	parsed: ParsedToken,
	index: number,
	lastValue: number[],
	lastVelocity: number[],
): number {
	// Missing slots read as NaN, which is what `undefined` arithmetic produced.
	const value = lastValue[index] ?? NaN;
	let next: number;
	if (parsed.repeat) {
		next = parsed.mode === 'explicit' ? value : value + (lastVelocity[index] ?? 0);
		lastValue[index] = next;
		return next;
	}
	if (parsed.mode === 'single') {
		lastVelocity[index] = parsed.value;
		next = value + parsed.value;
	} else if (parsed.mode === 'double') {
		const velocity = (lastVelocity[index] ?? NaN) + parsed.value;
		lastVelocity[index] = velocity;
		next = value + velocity;
	} else {
		next = parsed.value;
		lastVelocity[index] = 0;
	}
	lastValue[index] = next;
	return next;
}
