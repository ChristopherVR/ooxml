import {
	finiteFormulaValue as finite,
	compatibleFormulaUnits as compatible,
	formulaProductUnit,
	formulaQuotientUnit,
	formulaExponentUnit,
} from './formula-arithmetic';
import type { ShapeFrame } from './formula-extra-functions';
import { formulaFailure } from './formula';
import type { VisioFormulaAst, VisioFormulaReference, VisioFormulaValue } from './formula';

export function binaryOperator(
	operator: string,
	a: VisioFormulaValue,
	b: VisioFormulaValue,
): VisioFormulaValue {
	let unit: VisioFormulaValue['unit'];
	if (operator === '+' || operator === '-') {
		unit = compatible(a, b);
		return finite({ value: operator === '+' ? a.value + b.value : a.value - b.value, unit });
	}
	if (operator === '*') {
		return finite({ value: a.value * b.value, unit: formulaProductUnit(a.unit, b.unit) });
	}
	if (operator === '/') {
		if (b.value === 0) return formulaFailure('value', 'Division by zero');
		return finite({ value: a.value / b.value, unit: formulaQuotientUnit(a.unit, b.unit) });
	}
	if (operator === '^') {
		return finite({ value: a.value ** b.value, unit: formulaExponentUnit(a.unit, b) });
	}
	compatible(a, b);
	const comparisons: Record<string, boolean> = {
		'=': a.value === b.value,
		'<>': a.value !== b.value,
		'!=': a.value !== b.value,
		'<': a.value < b.value,
		'>': a.value > b.value,
		'<=': a.value <= b.value,
		'>=': a.value >= b.value,
	};
	if (!(operator in comparisons))
		return formulaFailure('syntax', `Unsupported operator ${operator}`);
	return { value: comparisons[operator] ? 1 : 0, unit: 'scalar' };
}

/** The frame of the formula's own shape (SHEETREF()), unflipped; other sheets are unsupported. */
export function evaluatedShapeFrame(
	sheet: VisioFormulaAst,
	resolve: (reference: VisioFormulaReference) => VisioFormulaValue,
): ShapeFrame {
	if (sheet.kind !== 'call' || sheet.name !== 'SHEETREF' || sheet.args.length)
		return formulaFailure('unsupported', 'A sheet other than SHEETREF()');
	const read = (cell: string, fallback?: number): number => {
		try {
			return finite(resolve({ cell })).value;
		} catch (error) {
			if (fallback === undefined) throw error;
			return fallback;
		}
	};
	if (read('FlipX', 0) !== 0 || read('FlipY', 0) !== 0)
		return formulaFailure('unsupported', 'A coordinate conversion on a flipped shape');
	return {
		pinX: read('PinX'),
		pinY: read('PinY'),
		locPinX: read('LocPinX', 0),
		locPinY: read('LocPinY', 0),
		angle: read('Angle', 0),
	};
}
