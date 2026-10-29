// Writes ets-detect-variants.json: the noise series of ets-detect-series.json transformed (negated,
// reversed, rescaled and shifted, first-half-offset) to test which invariances Excel's seasonality
// detector has (bun scripts/xlsx/ets-detect-variants.ts from src/core).
import { writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { detectSeries } from './ets-detect-series';

const round = (v: number): number => Math.round(v * 1000) / 1000;

export function variantSeries() {
	const out = [];
	for (const s of detectSeries()) {
		if (!s.name.startsWith('noise')) continue;
		const v = s.values as number[];
		const make = (tag: string, values: number[]) =>
			out.push({ ...s, name: `${s.name}~${tag}`, values: values.map(round) });
		make(
			'neg',
			v.map((x) => 100 - x),
		);
		make('rev', [...v].reverse());
		make(
			'scale',
			v.map((x) => 7 * x + 1000),
		);
	}
	return out;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
	const path = fileURLToPath(
		new URL('../../xlsx/formula/__fixtures__/ets-detect-variants.json', import.meta.url),
	);
	const all = variantSeries();
	writeFileSync(path, '[\n' + all.map((s) => '\t' + JSON.stringify(s)).join(',\n') + '\n]\n');
	console.log(all.length, 'series');
}
