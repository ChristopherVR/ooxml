import { formatVisioFieldValue, type VisioFieldValue } from './text-fields';
import { fractionalFieldNumber } from './text-field-fraction';

/** Native FIELDPICTURE numeric formats use the Value cell's display unit. */
export function numericFieldPicture(
	value: VisioFieldValue,
	code: number,
	displayUnit: string | undefined,
): string | undefined {
	if (
		value.kind !== 'number' ||
		code < 0 ||
		code > 18 ||
		(code > 9 && code < 15) ||
		!Number.isInteger(code)
	)
		return undefined;
	const unit = displayUnit?.toUpperCase() ?? 'NUM';
	const units: Record<string, [typeof value.unit, number, string]> = {
		NUM: ['scalar', 1, ''],
		IN: ['length', 1, ' in.'],
		MM: ['length', 25.4, ' mm.'],
		CM: ['length', 2.54, ' cm.'],
		FT: ['length', 1 / 12, ' ft.'],
		DEG: ['angle', 180 / Math.PI, ' deg.'],
		RAD: ['angle', 1, ' rad.'],
	};
	const definition = Object.hasOwn(units, unit) ? units[unit] : undefined;
	// Page-dependent units, compound feet/inches, dates and locale pictures remain cached.
	if (!definition || definition[0] !== value.unit) return undefined;
	const general = code < 2;
	const decimals = general ? 4 : Math.floor((code - 2) / 2);
	const picture = general ? '0.####' : decimals ? `0.${'0'.repeat(decimals)}` : '0';
	const displayed = value.value * definition[1];
	if (!Number.isFinite(displayed)) return undefined;
	if (code >= 15) {
		const fraction = fractionalFieldNumber(displayed, code < 17 ? 9 : 99);
		return fraction === undefined ? undefined : fraction + (code % 2 === 0 ? definition[2] : '');
	}
	// Integer overflow and larger general-number precision remain unverified.
	if ((general || !decimals) && Math.abs(displayed) > 1e15) return undefined;
	const scientific =
		decimals > 0 &&
		(Math.abs(displayed) >= 1e15 ||
			(displayed !== 0 && Math.abs(displayed) < 0.5 / 10 ** decimals));
	let text: string | undefined;
	if (scientific) {
		const [coefficient, power] = displayed.toExponential().split('e');
		let exponent = Number(power);
		let precision = Math.abs(displayed) >= 1e15 ? decimals : decimals - 1;
		const mantissaPicture = precision ? `0.${'0'.repeat(precision)}` : '0';
		let mantissa = formatVisioFieldValue(
			{ kind: 'number', value: Number(coefficient), unit: 'scalar' },
			mantissaPicture,
		);
		if (mantissa === undefined) return undefined;
		// Native Visio expands the fixed mantissa precision when rounding carries its exponent.
		if (Math.abs(Number(mantissa)) >= 10) {
			exponent += 1;
			precision = decimals;
			mantissa = Math.sign(displayed).toFixed(precision);
		}
		if (general) mantissa = mantissa.replace(/\.?0+$/, '');
		else if (!mantissa.includes('.')) mantissa += '.';
		text = mantissa + (code % 2 ? 'e' : 'E') + exponent;
	} else {
		text = formatVisioFieldValue({ kind: 'number', value: displayed, unit: 'scalar' }, picture);
	}
	return text === undefined ? undefined : text + (code % 2 ? definition[2] : '');
}
