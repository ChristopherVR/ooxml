import {
	compatibleFormulaUnits as compatible,
	finiteFormulaValue as finite,
} from './formula-arithmetic';
import { formulaFailure } from './formula';
import type { VisioFormulaAst, VisioFormulaValue } from './formula';

/** BOUND(value, type, ignore1, min1, max1, ...) over already-parsed arguments. */
export function evaluateBound(
	args: VisioFormulaAst[],
	arg: (index: number) => VisioFormulaValue,
): VisioFormulaValue {
	// BOUND(value, type, ignore1, min1, max1, ...): control handles of stencil shapes use it.
	// https://learn.microsoft.com/en-us/office/client-developer/visio/bound-function
	if (args.length < 5 || (args.length - 2) % 3 !== 0 || args.length > 302)
		return formulaFailure('arity', 'Invalid BOUND argument count');
	const value = arg(0),
		type = arg(1);
	if (type.unit !== 'scalar' || ![0, 1, 2].includes(type.value))
		return formulaFailure('value', 'BOUND type must be 0, 1 or 2');
	if (type.value === 2) return value;
	if (type.value === 1) {
		// Excluded ranges: a value strictly inside one moves to its nearest edge (the upper
		// one halfway), and may then fall inside another, so repeat.
		let current = value.value,
			unit = value.unit;
		for (let pass = 0; pass <= (args.length - 2) / 3; pass++) {
			let moved = false;
			for (let index = 2; index < args.length; index += 3) {
				if (arg(index).value !== 0) continue;
				const a = arg(index + 1),
					b = arg(index + 2);
				unit = compatible({ value: current, unit }, a);
				unit = compatible({ value: current, unit }, b);
				const low = Math.min(a.value, b.value),
					high = Math.max(a.value, b.value);
				if (current > low && current < high) {
					current = current - low < high - current ? low : high;
					moved = true;
				}
			}
			if (!moved) break;
		}
		return finite({ value: current, unit });
	}
	let nearest: number | undefined;
	let unit = value.unit;
	for (let index = 2; index < args.length; index += 3) {
		if (arg(index).value !== 0) continue;
		const a = arg(index + 1),
			b = arg(index + 2);
		unit = compatible({ value: value.value, unit }, a);
		unit = compatible({ value: value.value, unit }, b);
		const low = Math.min(a.value, b.value),
			high = Math.max(a.value, b.value);
		if (value.value >= low && value.value <= high) return finite({ value: value.value, unit });
		const edge = value.value < low ? low : high;
		if (nearest === undefined || Math.abs(edge - value.value) < Math.abs(nearest - value.value))
			nearest = edge;
	}
	return finite({ value: nearest ?? value.value, unit });
}
