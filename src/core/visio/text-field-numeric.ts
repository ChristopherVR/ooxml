import { formatVisioFieldValue, type VisioFieldValue } from './text-fields';

/** Native FIELDPICTURE decimal formats use the Value cell's display unit. */
export function numericFieldPicture(
	value: VisioFieldValue,
	code: number,
	displayUnit: string | undefined,
): string | undefined {
	if (value.kind !== 'number' || code < 2 || code > 9 || !Number.isInteger(code)) return undefined;
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
	// Page-dependent units, fractions, dates and locale pictures remain cached.
	if (!definition || definition[0] !== value.unit) return undefined;
	const decimals = Math.floor((code - 2) / 2);
	const picture = decimals ? `0.${'0'.repeat(decimals)}` : '0';
	const text = formatVisioFieldValue(
		{ kind: 'number', value: value.value * definition[1], unit: 'scalar' },
		picture,
	);
	return text === undefined ? undefined : text + (code % 2 ? definition[2] : '');
}
