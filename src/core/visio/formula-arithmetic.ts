import { formulaFailure } from './formula.js';
import type { VisioFormulaUnit, VisioFormulaValue } from './formula.js';

export function finiteFormulaValue(value: VisioFormulaValue): VisioFormulaValue {
	if (!Number.isFinite(value.value)) return formulaFailure('value', 'Formula result is not finite');
	if (!['scalar', 'length', 'area', 'angle', 'time'].includes(value.unit))
		return formulaFailure('unit', 'Unknown resolved unit');
	return value;
}
export function compatibleFormulaUnits(
	a: VisioFormulaValue,
	b: VisioFormulaValue,
): VisioFormulaUnit {
	if (a.unit === b.unit) return a.unit;
	if (a.unit === 'scalar' && a.value === 0) return b.unit;
	if (b.unit === 'scalar' && b.value === 0) return a.unit;
	return formulaFailure('unit', `Incompatible units ${a.unit} and ${b.unit}`);
}
// Only scalar, inches and square inches participate in compound dimensional algebra.
// Higher powers, mixed dimensions and inverse units deliberately remain unsupported.
const lengthPower = (unit: VisioFormulaUnit): number | undefined =>
	unit === 'scalar' ? 0 : unit === 'length' ? 1 : unit === 'area' ? 2 : undefined;
function fromPower(power: number): VisioFormulaUnit {
	if (power === 0) return 'scalar';
	if (power === 1) return 'length';
	if (power === 2) return 'area';
	return formulaFailure('unit', 'Unsupported dimensional power');
}
export function formulaProductUnit(a: VisioFormulaUnit, b: VisioFormulaUnit): VisioFormulaUnit {
	if (a === 'scalar') return b;
	if (b === 'scalar') return a;
	const left = lengthPower(a),
		right = lengthPower(b);
	if (left === undefined || right === undefined)
		return formulaFailure('unit', 'Mixed dimensions are unsupported');
	return fromPower(left + right);
}
export function formulaQuotientUnit(a: VisioFormulaUnit, b: VisioFormulaUnit): VisioFormulaUnit {
	if (b === 'scalar') return a;
	if (a === b) return 'scalar';
	const left = lengthPower(a),
		right = lengthPower(b);
	if (left === undefined || right === undefined)
		return formulaFailure('unit', 'Mixed dimensions are unsupported');
	return fromPower(left - right);
}
export function formulaExponentUnit(
	a: VisioFormulaUnit,
	exponent: VisioFormulaValue,
): VisioFormulaUnit {
	if (exponent.unit !== 'scalar') return formulaFailure('unit', 'Exponents must be scalar');
	if (a === 'scalar') return 'scalar';
	if (exponent.value === 0) return 'scalar';
	if (exponent.value === 1) return a;
	const power = lengthPower(a);
	if (power === undefined) return formulaFailure('unit', 'Unsupported dimensional power');
	return fromPower(power * exponent.value);
}
export function formulaSquareRootUnit(unit: VisioFormulaUnit): VisioFormulaUnit {
	if (unit === 'scalar') return 'scalar';
	if (unit === 'area') return 'length';
	return formulaFailure('unit', 'Square root requires scalar or square inches');
}
