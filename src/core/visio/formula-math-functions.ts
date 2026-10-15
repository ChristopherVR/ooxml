import { formulaFailure } from './formula';
import type { VisioFormulaValue } from './formula';
import { finiteFormulaValue, formulaExponentUnit } from './formula-arithmetic';

/**
 * Rounding, power, logarithm and inverse trigonometric functions of the ShapeSheet, the ones
 * shape formulas use most after arithmetic and IF (`ROUND(Width*0.5,2)`, `POW(Width,2)`,
 * `ATAN(Height/Width)`). Units follow the argument: rounding keeps the unit of the value, the
 * transcendental functions take and return scalars, the inverse trigonometric ones return angles.
 * https://learn.microsoft.com/en-us/office/client-developer/visio/round-function
 */
export const mathFormulaFunctions: ReadonlySet<string> = new Set([
	'ROUND',
	'INT',
	'TRUNC',
	'POW',
	'EXP',
	'LN',
	'LOG10',
	'ASIN',
	'ACOS',
	'ATAN',
]);

/** A quotient a rounding error away from a whole number is that whole number. */
function snap(value: number): number {
	const whole = Math.round(value);
	return Math.abs(value - whole) < 1e-9 * Math.max(1, Math.abs(value)) ? whole : value;
}

const scalar = (value: number): VisioFormulaValue => finiteFormulaValue({ value, unit: 'scalar' });

export function evaluateMathFormulaFunction(
	name: string,
	args: readonly VisioFormulaValue[],
): VisioFormulaValue {
	const arity = (min: number, max = min) => {
		if (args.length < min || args.length > max)
			return formulaFailure('arity', `Invalid ${name} argument count`);
	};
	const first = args[0];
	if (!first) return formulaFailure('arity', `Invalid ${name} argument count`);
	const value = finiteFormulaValue(first);
	if (name === 'ROUND') {
		arity(1, 2);
		const digits = args[1] ? finiteFormulaValue(args[1]) : { value: 0, unit: 'scalar' as const };
		if (digits.unit !== 'scalar' || !Number.isInteger(digits.value) || Math.abs(digits.value) > 15)
			return formulaFailure('value', 'ROUND requires a whole number of digits from -15 to 15');
		const factor = 10 ** digits.value;
		// Half away from zero, as the ShapeSheet shows it.
		let scaled = Math.abs(value.value) * factor;
		// A product a rounding error short of a half rounds up, as 1.005 does at two digits.
		const nearest = Math.round(scaled * 2) / 2;
		if (Math.abs(scaled - nearest) < 1e-9 * Math.max(1, scaled)) scaled = nearest;
		return finiteFormulaValue({
			value: (Math.sign(value.value) * Math.round(scaled)) / factor,
			unit: value.unit,
		});
	}
	if (name === 'INT' || name === 'TRUNC') {
		arity(1);
		const whole = snap(value.value);
		return finiteFormulaValue({
			value: name === 'INT' ? Math.floor(whole) : Math.trunc(whole),
			unit: value.unit,
		});
	}
	if (name === 'POW') {
		arity(2);
		const exponent = finiteFormulaValue(args[1]!);
		const result = value.value ** exponent.value;
		return finiteFormulaValue({ value: result, unit: formulaExponentUnit(value.unit, exponent) });
	}
	arity(1);
	if (value.unit !== 'scalar') return formulaFailure('unit', `${name} requires a scalar`);
	if (name === 'EXP') return scalar(Math.exp(value.value));
	if (name === 'LN' || name === 'LOG10') {
		if (!(value.value > 0)) return formulaFailure('value', `${name} requires a positive number`);
		return scalar(name === 'LN' ? Math.log(value.value) : Math.log10(value.value));
	}
	if (name === 'ATAN') return finiteFormulaValue({ value: Math.atan(value.value), unit: 'angle' });
	if (name === 'ASIN' || name === 'ACOS') {
		if (Math.abs(value.value) > 1)
			return formulaFailure('value', `${name} requires a number from -1 to 1`);
		return finiteFormulaValue({
			value: name === 'ASIN' ? Math.asin(value.value) : Math.acos(value.value),
			unit: 'angle',
		});
	}
	return formulaFailure('unsupported', `Unsupported function ${name}`);
}
