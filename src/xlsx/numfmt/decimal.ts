/**
 * Exact decimal digit handling for display. Excel keeps 15 significant digits, so a value is first
 * reduced to `toPrecision(15)` and then rounded half away from zero on its decimal digits; this
 * avoids binary artefacts such as `1.005` displaying as `1.00`.
 */
export interface Digits {
	/** Significant digits without leading or trailing zeros (`''` for zero). */
	d: string;
	/** Count of digits before the decimal point: value = 0.d x 10^point. */
	point: number;
}

/** Decimal digits of `|x|` at 15 significant digits. */
export function toDigits(x: number): Digits {
	const abs = Math.abs(x);
	if (abs === 0 || !Number.isFinite(abs)) return { d: '', point: 0 };
	const [mant = '0', expPart = '0'] = abs.toExponential(14).split('e');
	const digits = mant.replace('.', '');
	return trim({ d: digits, point: Number(expPart) + 1 });
}

function trim(v: Digits): Digits {
	let d = v.d;
	let point = v.point;
	let lead = 0;
	while (lead < d.length && d.charAt(lead) === '0') lead++;
	d = d.slice(lead);
	point -= lead;
	d = d.replace(/0+$/, '');
	return d ? { d, point } : { d: '', point: 0 };
}

/** Rounds to `keep` significant digits (half away from zero). */
export function roundSignificant(v: Digits, keep: number): Digits {
	if (keep >= v.d.length) return v;
	if (keep < 0) return { d: '', point: 0 };
	const roundUp = v.d.charAt(keep) >= '5';
	let head = v.d.slice(0, keep);
	let point = v.point;
	if (roundUp) {
		const chars = head.split('');
		let i = chars.length - 1;
		while (i >= 0 && chars[i] === '9') {
			chars[i] = '0';
			i--;
		}
		if (i < 0) {
			chars.unshift('1');
			point++;
		} else {
			chars[i] = String(Number(chars[i]) + 1);
		}
		head = chars.join('');
	}
	return trim({ d: head, point });
}

/** Rounds to `decimals` places after the point. */
export function roundDecimals(v: Digits, decimals: number): Digits {
	return roundSignificant(v, v.point + decimals);
}

/** Multiplies by 10^n exactly. */
export function shift(v: Digits, n: number): Digits {
	return v.d ? { d: v.d, point: v.point + n } : v;
}

/** The integer digits (no leading zeros, `''` when zero) of a value. */
export function integerPart(v: Digits): string {
	if (v.point <= 0) return '';
	return v.d.slice(0, v.point).padEnd(v.point, '0');
}

/** Exactly `decimals` fractional digits (truncated; round first). */
export function fractionPart(v: Digits, decimals: number): string {
	if (decimals <= 0) return '';
	let frac: string;
	if (v.point >= 0) frac = v.d.slice(v.point);
	else frac = '0'.repeat(-v.point) + v.d;
	return frac.slice(0, decimals).padEnd(decimals, '0');
}

export const isZero = (v: Digits): boolean => v.d === '';

/** Plain decimal string of a non-negative value, trailing zeros removed. */
export function plain(v: Digits): string {
	if (!v.d) return '0';
	const int = integerPart(v) || '0';
	const fracLen = Math.max(0, v.d.length - v.point);
	const frac = fractionPart(v, fracLen);
	return frac ? `${int}.${frac}` : int;
}
