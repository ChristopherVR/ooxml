import { ARRAY_FUNCTIONS } from './array.js';
import { DATETIME_FUNCTIONS } from './datetime.js';
import { ENGINEERING_FUNCTIONS } from './engineering.js';
import { FINANCIAL_FUNCTIONS } from './financial.js';
import { INFO_FUNCTIONS } from './info.js';
import { LOGICAL_FUNCTIONS } from './logical.js';
import { LOOKUP_FUNCTIONS } from './lookup.js';
import { MATH_FUNCTIONS } from './math.js';
import { MATRIX_FUNCTIONS } from './matrix.js';
import { STATISTICAL_FUNCTIONS } from './statistical.js';
import { DISTRIBUTION_FUNCTIONS } from './distributions.js';
import { TEXT_FUNCTIONS } from './text.js';
import { TEXT_EXTRA_FUNCTIONS } from './text-extra.js';
import type { FunctionInfo, FunctionSpec } from './types.js';

const ALL: readonly FunctionSpec[] = [
	...MATH_FUNCTIONS,
	...MATRIX_FUNCTIONS,
	...STATISTICAL_FUNCTIONS,
	...DISTRIBUTION_FUNCTIONS,
	...LOGICAL_FUNCTIONS,
	...TEXT_FUNCTIONS,
	...TEXT_EXTRA_FUNCTIONS,
	...DATETIME_FUNCTIONS,
	...LOOKUP_FUNCTIONS,
	...ARRAY_FUNCTIONS,
	...INFO_FUNCTIONS,
	...FINANCIAL_FUNCTIONS,
	...ENGINEERING_FUNCTIONS,
];

const REGISTRY = new Map<string, FunctionSpec>();
for (const fn of ALL) REGISTRY.set(fn.name, fn);

/** The implementation of a function by upper-case name (prefixes already stripped). */
export const getFunction = (name: string): FunctionSpec | undefined => REGISTRY.get(name);

/** Whether a function name (any case, with or without `_xlfn.`) is implemented. */
export const isVolatileFunction = (name: string): boolean => REGISTRY.get(name)?.volatile === true;

/** Every implemented function, sorted by name. */
export const FUNCTION_CATALOG: readonly FunctionInfo[] = Object.freeze(
	[...REGISTRY.values()]
		.map(({ name, category, syntax, description }) => ({ name, category, syntax, description }))
		.sort((a, b) => a.name.localeCompare(b.name)),
);
