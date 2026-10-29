// Writes probe series that rescale, shift or reverse series already recorded in ets-search-series.json
// to see whether Excel's FORECAST.ETS parameter search depends on the scale (an absolute tolerance)
// or the direction of the data:
//   bun scripts/xlsx/ets-scale-series.ts          (from src/core)
// then record Excel's answers with `xlsx/formula/__fixtures__/ets-search-cases.ps1 -In ets-scale-series.json`.
import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import type { EtsCaseInput } from './ets-series';

const fixtures = new URL('../../xlsx/formula/__fixtures__/', import.meta.url);
const base = JSON.parse(
	readFileSync(fileURLToPath(new URL('ets-search-series.json', fixtures)), 'utf8'),
) as EtsCaseInput[];

const SCALES = [1e-4, 1e-2, 1, 100, 1e4, 1e7];

function calls(n: number): string[] {
	return [
		'FORECAST.ETS.STAT({V},{L},1,0)',
		'FORECAST.ETS.STAT({V},{L},2,0)',
		'FORECAST.ETS.STAT({V},{L},7,0)',
		`FORECAST.ETS(${n + 1},{V},{L},0)`,
	];
}

const out: EtsCaseInput[] = [];
// Every 10th probe of the 8 to 40 point series: scales, a reversal and a negation.
base
	.filter((_c, i) => i % 10 === 3)
	.forEach((c) => {
		const n = c.values.length;
		const timeline = Array.from({ length: n }, (_x, k) => k + 1);
		for (const s of SCALES)
			out.push({
				name: `${c.name}|x${s}`,
				values: c.values.map((v) => (v === null ? v : v * s)),
				timeline,
				formulas: calls(n),
			});
		out.push({
			name: `${c.name}|rev`,
			values: [...c.values].reverse(),
			timeline,
			formulas: calls(n),
		});
		out.push({
			name: `${c.name}|neg`,
			values: c.values.map((v) => (v === null ? v : -v)),
			timeline,
			formulas: calls(n),
		});
	});
writeFileSync(
	fileURLToPath(new URL('ets-scale-series.json', fixtures)),
	'[\n' + out.map((r) => '\t' + JSON.stringify(r)).join(',\n') + '\n]\n',
);
console.log(out.length, 'series');
