// Writes the cell-floats documents (`docx/__fixtures__/cell-floats/cases.ts`) as .docx files for
// `docx-word-cell-float-probe.ps1` to render with Word.
//
//   bun scripts/docx-word-cell-float-cases.ts <out dir> [demo]
//
// With `demo`, writes only the combined demonstration document instead (cell-floats-demo.docx).
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { strToU8, unzipSync, zipSync, strFromU8 } from 'fflate';
import {
	CHART_PARTS,
	cellFloatCases,
	cellFloatDemo,
	cellFloatPackage,
} from '../docx/__fixtures__/cell-floats/cases';

const out = process.argv[2];
if (!out) throw new Error('usage: bun scripts/docx-word-cell-float-cases.ts <out dir> [demo]');
mkdirSync(out, { recursive: true });
const chart = unzipSync(readFileSync(new URL('../docx/__fixtures__/chart.docx', import.meta.url)));
const chartParts = Object.fromEntries(
	CHART_PARTS.map((part) => [part, strFromU8(chart[part] ?? new Uint8Array())]),
);
const write = (name: string, parts: Record<string, string>) =>
	writeFileSync(
		join(out, `${name}.docx`),
		zipSync(Object.fromEntries(Object.entries(parts).map(([path, xml]) => [path, strToU8(xml)]))),
	);
if (process.argv[3] === 'demo')
	write('cell-floats-demo', cellFloatPackage(cellFloatDemo(), chartParts));
else
	for (const [name, entry] of Object.entries(cellFloatCases()))
		write(name, cellFloatPackage(entry, chartParts));
