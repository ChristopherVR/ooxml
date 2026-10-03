import type { DigitChar, Token } from './types.js';

type FracToken =
	| { r: 'lit'; v: string }
	| { r: 'int'; index: number }
	| { r: 'num'; index: number }
	| { r: 'slash' }
	| { r: 'den'; index: number };

export interface FractionPlan {
	tokens: FracToken[];
	intDigits: DigitChar[];
	numDigits: DigitChar[];
	/** Denominator placeholders (empty when the denominator is fixed). */
	denDigits: DigitChar[];
	fixedDenominator?: number;
	/** Characters the denominator occupies (used to blank a zero fraction). */
	denWidth: number;
}

const isDigit = (t: Token | undefined): t is { t: 'digit'; c: DigitChar } => t?.t === 'digit';

/** Index of the fraction slash (a `/` between digit placeholders), or -1. */
export function fractionSlash(tokens: readonly Token[]): number {
	return tokens.findIndex(
		(t, i) =>
			t.t === 'slash' &&
			isDigit(tokens[i - 1]) &&
			(isDigit(tokens[i + 1]) || tokens[i + 1]?.t === 'num'),
	);
}

export function compileFraction(tokens: readonly Token[], slash: number): FractionPlan {
	let numStart = slash;
	while (isDigit(tokens[numStart - 1])) numStart--;
	let denEnd = slash + 1;
	while (isDigit(tokens[denEnd]) || tokens[denEnd]?.t === 'num') denEnd++;
	const denTokens = tokens.slice(slash + 1, denEnd);
	const fixed =
		denTokens.some((t) => t.t === 'num') &&
		denTokens.every((t) => t.t === 'num' || (isDigit(t) && t.c === '0'));
	const plan: FractionPlan = {
		tokens: [],
		intDigits: [],
		numDigits: [],
		denDigits: [],
		denWidth: denTokens.length,
	};
	if (fixed) {
		plan.fixedDenominator = Number(denTokens.map((t) => (t.t === 'num' ? t.v : '0')).join(''));
	}
	tokens.forEach((tok, i) => {
		if (i === slash) plan.tokens.push({ r: 'slash' });
		else if (i > slash && i < denEnd) {
			if (fixed) {
				if (i === slash + 1) plan.tokens.push({ r: 'den', index: 0 });
			} else if (isDigit(tok)) {
				plan.tokens.push({ r: 'den', index: plan.denDigits.length });
				plan.denDigits.push(tok.c);
			}
		} else if (isDigit(tok) && i >= numStart && i < slash) {
			plan.tokens.push({ r: 'num', index: plan.numDigits.length });
			plan.numDigits.push(tok.c);
		} else if (isDigit(tok) && i < numStart) {
			plan.tokens.push({ r: 'int', index: plan.intDigits.length });
			plan.intDigits.push(tok.c);
		} else if (tok.t === 'lit') plan.tokens.push({ r: 'lit', v: tok.v });
		else if (tok.t === 'num') plan.tokens.push({ r: 'lit', v: tok.v });
		else if (tok.t === 'digit') plan.tokens.push({ r: 'lit', v: tok.c });
		else if (tok.t === 'point') plan.tokens.push({ r: 'lit', v: '.' });
		else if (tok.t === 'slash') plan.tokens.push({ r: 'lit', v: '/' });
		else if (tok.t === 'percent') plan.tokens.push({ r: 'lit', v: '%' });
	});
	return plan;
}

/**
 * Rational approximation of `x` with a denominator of at most `maxDen`, as Excel computes it: the
 * last continued-fraction convergent whose denominator fits (semiconvergents are never used, so
 * 0.3 with `?/?` is 1/3 and 0.06 is 0/1).
 */
export function approximate(x: number, maxDen: number): [number, number] {
	if (Number.isInteger(x)) return [x, 1];
	let p0 = 0;
	let q0 = 1;
	let p1 = 1;
	let q1 = 0;
	let y = x;
	for (let iter = 0; iter < 64; iter++) {
		const a = Math.floor(y);
		const p2 = a * p1 + p0;
		const q2 = a * q1 + q0;
		if (q2 > maxDen) break;
		p0 = p1;
		q0 = q1;
		p1 = p2;
		q1 = q2;
		const rest = y - a;
		if (rest < 1e-10) break;
		y = 1 / rest;
	}
	return [p1, q1];
}

const pad = (c: DigitChar): string => (c === '0' ? '0' : c === '?' ? ' ' : '');

function padLeft(s: string, slots: readonly DigitChar[]): string {
	let prefix = '';
	for (let i = 0; i < slots.length - s.length; i++) prefix += pad(slots[i] ?? '#');
	return prefix + s;
}

function padRight(s: string, slots: readonly DigitChar[]): string {
	let suffix = '';
	for (let i = s.length; i < slots.length; i++) suffix += pad(slots[i] ?? '#');
	return s + suffix;
}

/** Formats `|value|` as a fraction. */
export function renderFraction(value: number, plan: FractionPlan): { text: string; zero: boolean } {
	const x = Math.abs(value);
	const hasInt = plan.intDigits.length > 0;
	let ip = hasInt ? Math.floor(x) : 0;
	const f = x - ip;
	const maxDen = 10 ** Math.max(1, plan.denDigits.length) - 1;
	let [n, d] = plan.fixedDenominator
		? [Math.round(f * plan.fixedDenominator), plan.fixedDenominator]
		: approximate(f, maxDen);
	if (hasInt && n >= d) {
		ip += 1;
		n = 0;
	}
	// A sign is kept for a mixed fraction that rounds to zero ("-0"), not for a bare "0/1".
	const zero = !hasInt && n === 0;
	const blank = hasInt && n === 0;
	const intText = hasInt ? padLeft(ip > 0 ? String(ip) : n === 0 ? '0' : '', plan.intDigits) : '';
	const allHash = (slots: readonly DigitChar[]): boolean => slots.every((c) => c === '#');
	const numHash = allHash(plan.numDigits);
	// With `#` numerator digits Excel drops the separator after an empty integer ("1/2" for
	// `# #/#`) and the whole fraction, separator included, when it is zero ("1" for `# #/#`).
	const hideFraction =
		blank && numHash && (plan.fixedDenominator !== undefined || allHash(plan.denDigits));
	const dropSeparator = hideFraction || (hasInt && intText === '' && numHash);
	const lastInt = plan.tokens.reduce((last, t, i) => (t.r === 'int' ? i : last), -1);
	const firstNum = plan.tokens.findIndex((t) => t.r === 'num');
	let text = '';
	let intDone = false;
	plan.tokens.forEach((tok, i) => {
		if (tok.r === 'lit') {
			if (!(dropSeparator && i > lastInt && i < firstNum)) text += tok.v;
		} else if (tok.r === 'int') {
			if (!intDone) text += intText;
			intDone = true;
		} else if (hideFraction) {
			// nothing
		} else if (blank) {
			if (tok.r === 'num') text += ' ';
			else if (tok.r === 'slash') text += ' ';
			else text += ' '.repeat(plan.fixedDenominator ? plan.denWidth : 1);
		} else if (tok.r === 'num') {
			if (tok.index === 0) text += padLeft(String(n), plan.numDigits);
		} else if (tok.r === 'slash') text += '/';
		else if (plan.fixedDenominator) text += String(d);
		else if (tok.index === 0) text += padRight(String(d), plan.denDigits);
	});
	return { text, zero };
}
