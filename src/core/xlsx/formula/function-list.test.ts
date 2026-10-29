// The committed function list (viewers/xlsx/docs/functions.md) must match the registry.
// Regenerate with `bun scripts/xlsx/function-list.ts` from src/core.
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import {
	buildFunctionList,
	DEVIATIONS,
	excelVerifiedNames,
	FUNCTION_LIST_PATH,
} from '../../scripts/xlsx/function-list';
import { FUNCTION_CATALOG } from './functions/registry';

describe('supported function list', () => {
	it('matches the committed viewers/xlsx/docs/functions.md', () => {
		const committed = readFileSync(FUNCTION_LIST_PATH, 'utf8').replace(/\r\n/g, '\n');
		expect(committed, 'run `bun scripts/xlsx/function-list.ts` in src/core').toBe(
			buildFunctionList(),
		);
	});

	it('only annotates functions the registry implements', () => {
		const names = new Set(FUNCTION_CATALOG.map((f) => f.name));
		for (const name of Object.keys(DEVIATIONS)) expect(names.has(name), name).toBe(true);
	});

	it('marks functions with recorded Excel results', () => {
		const verified = excelVerifiedNames();
		for (const name of ['SUM', 'DDB', 'AMORLINC', 'XMATCH'])
			expect(verified.has(name), name).toBe(true);
	});
});
