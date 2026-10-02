import type { CallContext, LazyArg } from '../context.js';
import type { Value } from '../values.js';

/**
 * `value` parameters take one value: a multi-cell range or array lifts the call over its
 * elements (dynamic arrays); `any` parameters receive references and arrays unchanged.
 */
export type ParamKind = 'value' | 'any';

export interface FunctionSpec {
	name: string;
	category: string;
	syntax: string;
	description: string;
	minArgs: number;
	maxArgs: number;
	/** Per position; the last entry repeats. Defaults to every parameter `value`. */
	params?: readonly ParamKind[];
	/** Recalculated on every recalculation (NOW, RAND, OFFSET, INDIRECT, ...). */
	volatile?: boolean;
	fn?: (args: Value[], ctx: CallContext) => Value;
	lazy?: (args: LazyArg[], ctx: CallContext) => Value;
}

/** Public catalog entry for an implemented function (function wizard, autocomplete). */
export interface FunctionInfo {
	name: string;
	category: string;
	syntax: string;
	description: string;
}

export const paramKind = (spec: FunctionSpec, index: number): ParamKind => {
	const params = spec.params;
	if (!params || params.length === 0) return 'value';
	return params[Math.min(index, params.length - 1)] ?? 'value';
};
