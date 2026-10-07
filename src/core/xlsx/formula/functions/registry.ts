import { ARRAY_FUNCTIONS } from './array';
import { DATABASE_FUNCTIONS } from './database';
import { DATETIME_FUNCTIONS } from './datetime';
import { ENGINEERING_FUNCTIONS } from './engineering';
import { FINANCIAL_FUNCTIONS } from './financial';
import { INFO_FUNCTIONS } from './info';
import { LOGICAL_FUNCTIONS } from './logical';
import { LOOKUP_FUNCTIONS } from './lookup';
import { MATH_FUNCTIONS } from './math';
import { MATRIX_FUNCTIONS } from './matrix';
import { STATISTICAL_FUNCTIONS } from './statistical';
import { DISTRIBUTION_FUNCTIONS } from './distributions';
import { TEXT_FUNCTIONS } from './text';
import { TEXT_EXTRA_FUNCTIONS } from './text-extra';
import type { FunctionInfo, FunctionSpec } from './types';

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
	...DATABASE_FUNCTIONS,
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
