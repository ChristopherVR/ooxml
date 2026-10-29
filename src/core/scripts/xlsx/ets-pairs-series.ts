// Writes probe series for the (alpha, beta) pairs of Excel's FORECAST.ETS search:
//   bun scripts/xlsx/ets-pairs-series.ts          (from src/core)
// then record Excel's answers with `xlsx/formula/__fixtures__/ets-search-cases.ps1 -In
// ets-pairs-series.json -Out ...`. Random series are drawn and kept when the lowest hold-out cost
// over the candidate pairs (`ets-search-map.ts`) is at a pair we want to pin, by a clear margin over
// the next pair: `proxy` keys (is the pair behind the outcome the proxy or another one?) and the
// large-beta keys with a spread of margins, lengths and trends (what separates the drift forecast
// from the pair forecast?). The committed ets-pairs-series.json was drawn with the candidate list of
// that time (with stand-in pairs), so running this again draws other series.
import { writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { etsHoldout, etsRun } from '../../xlsx/formula/functions/ets-core';
import { initialStates } from '../../xlsx/formula/functions/ets-init';
import { CANDIDATES } from '../../xlsx/formula/functions/ets-search-map';
import { etsSearchSeries, etsValidationCost } from '../../xlsx/formula/functions/ets-validate';
import type { EtsCaseInput } from './ets-series';
import { KINDS, make, rng, search } from './ets-search-series';

export interface Scored {
	key: string;
	margin: number;
	second: string;
}

/** The lowest-cost candidate of `searchLattice`, and the cost ratio of the next different pair. */
export function score(y: number[]): Scored {
	const n = y.length;
	const h = etsHoldout(n, 0);
	const ys = etsSearchSeries(y);
	const rows = CANDIDATES.map(([a, b]) => {
		const p = { alpha: a, beta: b, gamma: 0.001 };
		return {
			key: `${a}/${b}`,
			j: etsValidationCost(ys, 0, initialStates(y, 0, p, etsRun), p, h),
		};
	}).sort((x, z) => x.j - z.j);
	return { key: rows[0]!.key, margin: rows[1]!.j / rows[0]!.j, second: rows[1]!.key };
}

const WANT = [
	'0.9/0.099',
	'0.9/0.249',
	'0.9/0.499',
	'0.998/0.125',
	'0.998/0.099',
	'0.998/0.249',
	'0.998/0.499',
	'0.75/0.125',
	'0.25/0.249',
	'0.5/0.499',
	'0.5/0.25',
	'0.75/0.749',
	'0.126/0.125',
	'0.9/0.899',
	'0.666667/0.166667',
];

export function buildPairCases(perKey = 14): EtsCaseInput[] {
	const r = rng(777);
	const found = new Map<string, { y: number[]; s: Scored; kind: string }[]>();
	for (let i = 0; i < 60000; i++) {
		const kind = KINDS[i % KINDS.length]!;
		const n = [11, 12, 14, 16, 18, 20, 24, 28, 30, 36, 40][Math.floor(r() * 11)]!;
		const y = make(kind, n, r);
		const s = score(y);
		if (!WANT.includes(s.key)) continue;
		const list = found.get(s.key) ?? [];
		list.push({ y, s, kind });
		found.set(s.key, list);
	}
	const out: EtsCaseInput[] = [];
	for (const key of WANT) {
		const list = (found.get(key) ?? []).sort((a, b) => b.s.margin - a.s.margin);
		// Spread over the margins: the strongest few, then evenly through the rest.
		const pick = list.filter(
			(_v, i) => i < 4 || i % Math.max(1, Math.floor(list.length / (perKey - 4))) === 0,
		);
		pick
			.slice(0, perKey)
			.forEach((c, i) =>
				out.push(
					search(
						`pair-${key.replace('/', '_')}-${c.kind}-${c.y.length}-m${c.s.margin.toFixed(2)}-${i}`,
						c.y,
					),
				),
			);
	}
	return out;
}

if (import.meta.main) {
	const out = fileURLToPath(
		new URL('../../xlsx/formula/__fixtures__/ets-pairs-series.json', import.meta.url),
	);
	const cases = buildPairCases();
	writeFileSync(out, `[\n${cases.map((c) => `\t${JSON.stringify(c)}`).join(',\n')}\n]\n`);
	console.log(cases.length);
}
