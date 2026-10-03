import {
	type Digits,
	fractionPart,
	integerPart,
	isZero,
	roundDecimals,
	shift,
	toDigits,
} from './decimal.js';
import type { DigitChar, Token } from './types.js';

/** A number-section token with its role resolved at compile time. */
export type PlanToken =
	| { r: 'lit'; v: string }
	| { r: 'int'; c: DigitChar; index: number }
	| { r: 'point' }
	| { r: 'frac'; c: DigitChar; index: number }
	| { r: 'exp' };

export interface NumberPlan {
	tokens: PlanToken[];
	intDigits: DigitChar[];
	fracDigits: DigitChar[];
	grouping: boolean;
	/** Power of ten applied before formatting (`%` adds 2, a scaling comma subtracts 3). */
	shift: number;
	exponent?: { plus: boolean; digits: number };
}

const isDigit = (t: Token | undefined): t is { t: 'digit'; c: DigitChar } => t?.t === 'digit';

/** Resolves roles of a plain (non-fraction) number section. */
export function compileNumber(tokens: readonly Token[]): NumberPlan {
	const expAt = tokens.findIndex((t) => t.t === 'exp');
	const end = expAt < 0 ? tokens.length : expAt;
	let pointAt = tokens.findIndex((t, i) => t.t === 'point' && i < end);
	if (pointAt < 0) pointAt = end;
	const intDigits: DigitChar[] = [];
	const fracDigits: DigitChar[] = [];
	const out: PlanToken[] = [];
	let grouping = false;
	let scaleCommas = 0;
	let percent = 0;
	let exponent: NumberPlan['exponent'];
	const hasIntAfter = (k: number): boolean => {
		for (let j = k + 1; j < pointAt; j++) if (isDigit(tokens[j])) return true;
		return false;
	};
	let seenDigit = false;
	for (let i = 0; i < tokens.length; i++) {
		const tok = tokens[i];
		if (!tok) continue;
		if (i === expAt && tok.t === 'exp') {
			let n = 0;
			while (isDigit(tokens[i + 1 + n])) n++;
			exponent = { plus: tok.plus, digits: n };
			out.push({ r: 'exp' });
			i += n;
			continue;
		}
		switch (tok.t) {
			case 'digit':
				seenDigit = true;
				if (i < pointAt) {
					out.push({ r: 'int', c: tok.c, index: intDigits.length });
					intDigits.push(tok.c);
				} else if (i < end) {
					out.push({ r: 'frac', c: tok.c, index: fracDigits.length });
					fracDigits.push(tok.c);
				} else out.push({ r: 'lit', v: tok.c });
				break;
			case 'point':
				out.push(i === pointAt ? { r: 'point' } : { r: 'lit', v: '.' });
				break;
			case 'comma':
				if (i < pointAt && seenDigit && hasIntAfter(i)) grouping = true;
				else if (seenDigit && i < end) scaleCommas++;
				else out.push({ r: 'lit', v: ',' });
				break;
			case 'percent':
				percent++;
				out.push({ r: 'lit', v: '%' });
				break;
			case 'lit':
				out.push({ r: 'lit', v: tok.v });
				break;
			case 'num':
				out.push({ r: 'lit', v: tok.v });
				break;
			case 'slash':
				out.push({ r: 'lit', v: '/' });
				break;
			default:
				break;
		}
	}
	if (intDigits.length === 0 && fracDigits.length > 0) {
		const at = out.findIndex((t) => t.r === 'point');
		out.splice(at < 0 ? 0 : at, 0, { r: 'int', c: '#', index: 0 });
		intDigits.push('#');
	}
	const plan: NumberPlan = {
		tokens: out,
		intDigits,
		fracDigits,
		grouping,
		shift: 2 * percent - 3 * scaleCommas,
	};
	if (exponent) plan.exponent = exponent;
	return plan;
}

const pad = (c: DigitChar): string => (c === '0' ? '0' : c === '?' ? ' ' : '');

/** Distributes integer digits over the placeholders (extra digits go to the leftmost one). */
function placeIntegers(digits: string, slots: readonly DigitChar[], grouping: boolean): string[] {
	const n = slots.length;
	const out: string[] = slots.map((c, j) => {
		const p = n - 1 - j;
		return p < digits.length ? digits.charAt(digits.length - 1 - p) : pad(c);
	});
	if (n > 0 && digits.length > n) out[0] = digits.slice(0, digits.length - n) + (out[0] ?? '');
	if (!grouping) return out;
	let remaining = out.join('').replace(/[^0-9]/g, '').length;
	return out.map((s) => {
		let res = '';
		for (const ch of s) {
			res += ch;
			if (ch >= '0' && ch <= '9') {
				remaining--;
				if (remaining > 0 && remaining % 3 === 0) res += ',';
			}
		}
		return res;
	});
}

function placeFraction(digits: string, slots: readonly DigitChar[]): string[] {
	const out = slots.map((_, i) => digits.charAt(i));
	for (let i = slots.length - 1; i >= 0; i--) {
		const c = slots[i] ?? '0';
		if (digits.charAt(i) !== '0' || c === '0') break;
		out[i] = pad(c);
	}
	return out;
}

interface Mantissa {
	value: Digits;
	exp: number;
}

function scientific(v: Digits, plan: NumberPlan): Mantissa {
	if (isZero(v)) return { value: v, exp: 0 };
	const n = plan.intDigits.length;
	const engineering = n > 1 && plan.intDigits.some((c) => c !== '0');
	const period = engineering ? n : 1;
	const e = v.point - 1;
	let exp = Math.floor(e / period) * period;
	let m = roundDecimals(shift(v, -exp), plan.fracDigits.length);
	if (m.point > period) {
		exp += period;
		m = roundDecimals(shift(v, -exp), plan.fracDigits.length);
	}
	return { value: m, exp };
}

/** Formats `|value|`; `zero` reports whether the displayed number rounded to zero. */
export function renderNumber(value: number, plan: NumberPlan): { text: string; zero: boolean } {
	const v = shift(toDigits(value), plan.shift);
	let rounded: Digits;
	let exp = 0;
	if (plan.exponent) {
		const m = scientific(v, plan);
		rounded = m.value;
		exp = m.exp;
	} else rounded = roundDecimals(v, plan.fracDigits.length);
	// Excel fills every integer placeholder of a zero mantissa: 0 with `##0.0E+0` is 000.0E+0.
	const ints =
		plan.exponent && isZero(rounded)
			? plan.intDigits.map(() => '0')
			: placeIntegers(integerPart(rounded), plan.intDigits, plan.grouping);
	const fracs = placeFraction(fractionPart(rounded, plan.fracDigits.length), plan.fracDigits);
	let text = '';
	for (const tok of plan.tokens) {
		if (tok.r === 'lit') text += tok.v;
		else if (tok.r === 'int') text += ints[tok.index] ?? '';
		else if (tok.r === 'frac') text += fracs[tok.index] ?? '';
		else if (tok.r === 'point') text += '.';
		else if (plan.exponent) {
			const sign = exp < 0 ? '-' : plan.exponent.plus ? '+' : '';
			text += `E${sign}${String(Math.abs(exp)).padStart(plan.exponent.digits, '0')}`;
		}
	}
	return { text, zero: isZero(rounded) };
}
