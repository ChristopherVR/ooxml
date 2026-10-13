/** Nearest fraction with a one/two-digit denominator, as measured in native Visio. */
export function fractionalFieldNumber(value: number, maxDenominator: 9 | 99): string | undefined {
	if (!Number.isFinite(value) || Math.abs(value) > 1e15) return undefined;
	const absolute = Math.abs(value);
	let whole = Math.floor(absolute);
	const remainder = absolute - whole;
	let numerator = 0,
		denominator = 1,
		error = remainder;
	for (let candidate = 1; candidate <= maxDenominator; candidate++) {
		const top = Math.round(remainder * candidate);
		const fraction = top / candidate;
		const distance = Math.abs(remainder - fraction);
		if (distance < error || (distance === error && fraction > numerator / denominator)) {
			numerator = top;
			denominator = candidate;
			error = distance;
		}
	}
	if (numerator === denominator) {
		whole++;
		numerator = 0;
	}
	const text = numerator ? `${whole ? whole + ' ' : ''}${numerator}/${denominator}` : String(whole);
	return value < 0 && text !== '0' ? '-' + text : text;
}
