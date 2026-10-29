// Writes probe series of 8 to 10 points (no hold-out) for Excel's FORECAST.ETS search:
//   bun scripts/xlsx/ets-short-series.ts          (from src/core)
// then record Excel's answers with `xlsx/formula/__fixtures__/ets-search-cases.ps1 -In
// ets-short-series.json -Out ...`. 40 series of each length, every kind of `ets-search-series.ts`.
import { writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { KINDS, make, rng, search } from './ets-search-series';

const r = rng(4242);
const cases = [8, 9, 10].flatMap((n) =>
	Array.from({ length: 40 }, (_v, i) =>
		search(`short-${n}-${i}`, make(KINDS[i % KINDS.length]!, n, r)),
	),
);
const out = fileURLToPath(
	new URL('../../xlsx/formula/__fixtures__/ets-short-series.json', import.meta.url),
);
writeFileSync(out, `[\n${cases.map((c) => `\t${JSON.stringify(c)}`).join(',\n')}\n]\n`);
console.log(cases.length);
