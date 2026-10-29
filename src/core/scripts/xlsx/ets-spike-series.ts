// Writes probe series with a known, simple squared-error surface for Excel's FORECAST.ETS parameter
// search: a constant with one spike, and a constant with one level step, at every position:
//   bun scripts/xlsx/ets-spike-series.ts          (from src/core)
// then record Excel's answers with `xlsx/formula/__fixtures__/ets-search-cases.ps1 -In ets-spike-series.json`.
import { writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import type { EtsCaseInput } from './ets-series';

const out: EtsCaseInput[] = [];
for (const n of [8, 12, 20, 30]) {
	const timeline = Array.from({ length: n }, (_x, i) => i + 1);
	const formulas = [
		'FORECAST.ETS.STAT({V},{L},1,0)',
		'FORECAST.ETS.STAT({V},{L},2,0)',
		'FORECAST.ETS.STAT({V},{L},7,0)',
		`FORECAST.ETS(${n + 1},{V},{L},0)`,
	];
	for (let p = 0; p < n; p++) {
		out.push({
			name: `spike-${n}-${p}`,
			values: Array.from({ length: n }, (_x, i) => (i === p ? 60 : 50)),
			timeline,
			formulas,
		});
		if (p > 0)
			out.push({
				name: `step-${n}-${p}`,
				values: Array.from({ length: n }, (_x, i) => (i >= p ? 60 : 50)),
				timeline,
				formulas,
			});
	}
}
writeFileSync(
	fileURLToPath(new URL('../../xlsx/formula/__fixtures__/ets-spike-series.json', import.meta.url)),
	'[\n' + out.map((r) => '\t' + JSON.stringify(r)).join(',\n') + '\n]\n',
);
console.log(out.length, 'series');
