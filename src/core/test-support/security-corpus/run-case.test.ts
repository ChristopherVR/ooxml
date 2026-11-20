import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { afterAll, describe, expect, it } from 'vitest';

import { runCase } from './run-case';

const dir = mkdtempSync(join(tmpdir(), 'security-corpus-run-'));
afterAll(() => rmSync(dir, { recursive: true, force: true }));

const bundle = (name: string, source: string): string => {
	const file = join(dir, `${name}.mjs`);
	writeFileSync(file, source);
	return file;
};
const limits = { ms: 10_000, heapMb: 128 };

describe('runCase', () => {
	it('names the exit of a process that died before its case began, after trying again', async () => {
		const outcome = await runCase(bundle('dies', 'process.exit(7);'), 0, limits);
		expect(outcome).toMatchObject({ kind: 'crashed' });
		const error = (outcome as { error: string }).error;
		expect(error).toMatch(/never started the case \(3 tries\)/u);
		expect(error).toMatch(/exited before the case began, with exit code 7/u);
	});

	it('reports a crash while the case runs as the outcome, without retrying', async () => {
		const outcome = await runCase(
			bundle('crashes', 'export const cases = [{ run: () => process.exit(3) }];'),
			0,
			limits,
		);
		expect(outcome).toEqual({
			kind: 'crashed',
			error: 'exited while the case ran, with exit code 3',
		});
	});

	it('returns what a case that finishes reports', async () => {
		const outcome = await runCase(
			bundle('runs', 'export const cases = [{ run: async () => 42 }];'),
			0,
			limits,
		);
		expect(outcome).toMatchObject({ kind: 'done', value: 42, error: undefined, network: [] });
	});
});
