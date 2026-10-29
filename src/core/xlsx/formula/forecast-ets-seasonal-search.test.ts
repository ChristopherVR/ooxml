// The parameter search of seasonal FORECAST.ETS fits whose starting states are Excel's own
// (ets-init-chain.ts: 4 whole seasons or 5 and more), against the 700 such series of the start-state
// probes (__fixtures__/ets-init*-excel.json). Excel's triple is always one of 18 (alpha, beta, gamma)
// triples, and the lowest hold-out cost over them is its triple on 697 of the 700 series
// (functions/ets-search.ts). Before that lattice the search reproduced 427 of 700.
import { readdirSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { fitEts } from './functions/ets-core';
import { isExactLength } from './functions/ets-init-chain';

interface Probe {
	y: number[];
	m: number;
	alpha: number;
	beta: number;
	gamma: number;
}

function probes(): Probe[] {
	const dir = fileURLToPath(new URL('./__fixtures__/', import.meta.url));
	const out: Probe[] = [];
	for (const f of readdirSync(dir).filter((x) => /^ets-init\d?-excel\.json$/.test(x))) {
		const cases = JSON.parse(readFileSync(dir + f, 'utf8')) as {
			values?: (number | null)[];
			results?: { formula: string; type: string; value: number }[];
		}[];
		for (const c of cases) {
			if (!c.values || !c.results || c.values.some((v) => v === null)) continue;
			const stat = (k: number) =>
				c.results?.find((r) => r.formula.startsWith(`FORECAST.ETS.STAT({V},{L},${k},`));
			const a = stat(1);
			const m = Number(/,(\d+)\)$/.exec(a?.formula ?? '')?.[1]);
			if (!a || a.type !== 'number' || m < 2 || !isExactLength(c.values.length, m)) continue;
			out.push({
				y: c.values as number[],
				m,
				alpha: a.value,
				beta: (stat(2) as { value: number }).value,
				gamma: (stat(3) as { value: number }).value,
			});
		}
	}
	return out;
}

describe('seasonal parameter search with exact starting states', () => {
	const cases = probes();

	it('records 700 series', () => {
		expect(cases.length).toBe(700);
	});

	it('finds Excel alpha, beta and gamma together on 697 of them', () => {
		const same = cases.filter((c) => {
			const f = fitEts(c.y, c.m);
			return (
				Math.abs(f.alpha - c.alpha) < 1e-9 &&
				Math.abs(f.beta - c.beta) < 1e-9 &&
				Math.abs(f.gamma - c.gamma) < 1e-9
			);
		});
		expect(same.length).toBeGreaterThanOrEqual(697);
	});
});
