/** Function names accepted by Excel's Subtotal and Consolidate commands. */
export type SubtotalFunction =
	| 'average'
	| 'count'
	| 'counta'
	| 'max'
	| 'min'
	| 'product'
	| 'stdev'
	| 'stdevp'
	| 'sum'
	| 'var'
	| 'varp';

/** Excel SUBTOTAL function numbers (1-11). */
export type SubtotalFunctionCode = 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10 | 11;

export type SubtotalFunctionInput =
	| SubtotalFunction
	| Uppercase<SubtotalFunction>
	| SubtotalFunctionCode;

export interface SubtotalFunctionMeta {
	readonly code: SubtotalFunctionCode;
	readonly name: SubtotalFunction;
	readonly suffix: string;
	readonly grandLabel: string;
}

const FUNCTIONS: readonly SubtotalFunctionMeta[] = [
	{ code: 1, name: 'average', suffix: 'Average', grandLabel: 'Grand Average' },
	{ code: 2, name: 'count', suffix: 'Count', grandLabel: 'Grand Count' },
	{ code: 3, name: 'counta', suffix: 'Count', grandLabel: 'Grand Count' },
	{ code: 4, name: 'max', suffix: 'Max', grandLabel: 'Grand Max' },
	{ code: 5, name: 'min', suffix: 'Min', grandLabel: 'Grand Min' },
	{ code: 6, name: 'product', suffix: 'Product', grandLabel: 'Grand Product' },
	{ code: 7, name: 'stdev', suffix: 'StdDev', grandLabel: 'Grand StdDev' },
	{ code: 8, name: 'stdevp', suffix: 'StdDevp', grandLabel: 'Grand StdDevp' },
	{ code: 9, name: 'sum', suffix: 'Total', grandLabel: 'Grand Total' },
	{ code: 10, name: 'var', suffix: 'Var', grandLabel: 'Grand Var' },
	{ code: 11, name: 'varp', suffix: 'Varp', grandLabel: 'Grand Varp' },
];

const BY_CODE = new Map<number, SubtotalFunctionMeta>(FUNCTIONS.map((f) => [f.code, f]));
const BY_NAME = new Map<string, SubtotalFunctionMeta>(FUNCTIONS.map((f) => [f.name, f]));

/** Resolves a Subtotal / Consolidate function name or code (1-11) to its metadata. */
export function resolveSubtotalFunction(fn: SubtotalFunctionInput): SubtotalFunctionMeta {
	const meta = typeof fn === 'number' ? BY_CODE.get(fn) : BY_NAME.get(fn.trim().toLowerCase());
	if (!meta) throw new RangeError(`Unsupported subtotal function: ${String(fn)}`);
	return meta;
}
