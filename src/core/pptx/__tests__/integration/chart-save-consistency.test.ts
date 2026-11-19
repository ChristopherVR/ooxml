// Every chart save writes ranges and caches that agree with the embedded workbook, whatever the
// undo history kept and however much of it the saved-state budget and garbage collection left.
// CI runs a few cases that reach every fallback; `CHART_SAVE_SWEEP=full` runs every committed
// chart deck under every history copy, budget and collection timing, and
// `CHART_SAVE_SWEEP_OUT=<folder>` writes each saved deck there, for opening in PowerPoint, and a
// `summary.txt` with every case, its saves, issues and time.
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import { describe, expect, it } from 'vitest';

import {
	caseName,
	runSweepCase,
	sweepableCharts,
	wholeColumnChartDeck,
	type HistoryCopy,
	type SweepCase,
} from '../chart-save-sweep';
import { PptxHandler } from '../../index';
import type { ChartPptxElement } from '../../index';
import { requireFixture } from '../require-fixture';

const FULL = process.env['CHART_SAVE_SWEEP'] === 'full';
const OUT = process.env['CHART_SAVE_SWEEP_OUT'];

const fixture = (name: string) =>
	new Uint8Array(
		readFileSync(requireFixture(fileURLToPath(new URL(`../fixtures/${name}`, import.meta.url)))),
	);
const E2E_DECKS = [
	'chart-data-fidelity.pptx',
	'chart-ext-datalabels-range.pptx',
	'chart-ext-filtered-titles.pptx',
	'chart-filtered-series.pptx',
	'chart-label-callouts.pptx',
];

async function decks() {
	const blank = fixture('embedded-assets-sample.pptx');
	const list: { deck: string; bytes: Uint8Array; charts: string[] }[] = [];
	for (const categories of [true, false]) {
		const { bytes, chartPartPath } = await wholeColumnChartDeck(blank, categories);
		list.push({
			deck: categories ? 'whole-column' : 'whole-column-text',
			bytes,
			charts: [chartPartPath],
		});
	}
	for (const name of E2E_DECKS) {
		const bytes = fixture(`e2e/${name}`);
		list.push({ deck: name.replace(/\.pptx$/u, ''), bytes, charts: await sweepableCharts(bytes) });
	}
	return list;
}

type Variant = Pick<SweepCase, 'copy' | 'budget' | 'gc'>;

// CI: kept states replayed exactly, the JSON history at the default budget that lost its states
// to collection, and a budget that keeps none, on PowerPoint column and bubble charts and a
// whole-column chart.
const CI: Record<string, { charts: string[]; variants: Variant[] }> = {
	'chart-data-fidelity': {
		charts: ['ppt/charts/chart3.xml', 'ppt/charts/chart1.xml'],
		variants: [
			{ copy: 'clone', budget: undefined, gc: true },
			{ copy: 'json', budget: undefined, gc: true },
			{ copy: 'clone', budget: 0, gc: false },
		],
	},
	'whole-column': {
		charts: ['ppt/charts/whole.xml'],
		variants: [
			{ copy: 'json', budget: 0, gc: true },
			{ copy: 'clone', budget: undefined, gc: false },
		],
	},
};

function everyVariant(): Variant[] {
	const all: Variant[] = [];
	for (const copy of ['clone', 'structured', 'json'] as HistoryCopy[])
		for (const budget of [undefined, 0, 2048])
			for (const gc of [true, false]) all.push({ copy, budget, gc });
	return all;
}

async function cases(): Promise<SweepCase[]> {
	const all: SweepCase[] = [];
	for (const { deck, bytes, charts } of await decks()) {
		const ci = CI[deck];
		const chosen = FULL ? charts : charts.filter((chart) => ci?.charts.includes(chart));
		for (const chartPartPath of chosen)
			for (const variant of FULL ? everyVariant() : (ci?.variants ?? []))
				all.push({ deck, bytes, chartPartPath, ...variant });
	}
	return all;
}

describe('chart save consistency', () => {
	it(
		'saves ranges, caches and workbook cells that agree after every edit, undo and redo',
		async () => {
			const problems: string[] = [];
			const report: string[] = [];
			let saves = 0;
			for (const c of await cases()) {
				const name = caseName(c);
				const started = performance.now();
				const results = await runSweepCase(c);
				saves += results.length;
				report.push(
					`${name}: ${results.length} saves, ${results.flatMap((r) => r.issues).length} issues, ${results.flatMap((r) => r.base).length} base, ${Math.round(performance.now() - started)} ms`,
				);
				for (const { label, issues, bytes } of results) {
					for (const issue of issues)
						problems.push(`${name} ${label} ${issue.kind}: ${issue.detail}`);
					if (OUT) {
						mkdirSync(OUT, { recursive: true });
						writeFileSync(`${OUT}/${name.replace(/[/:]/gu, '_')}-${label}.pptx`, bytes);
					}
				}
			}
			report.push(`chart save sweep: ${saves} saves, ${problems.length} problems`, ...problems);
			if (OUT) writeFileSync(`${OUT}/summary.txt`, report.join('\n') + '\n');
			expect(saves).toBeGreaterThan(0);
			expect(problems).toEqual([]);
		},
		FULL ? 3_600_000 : 300_000,
	);
});

describe('category labels of a chart that names none', () => {
	it('are not read from the first column of its workbook', async () => {
		// Excel labels such a chart 1, 2, 3; here the first column holds a series of its own.
		const { bytes } = await wholeColumnChartDeck(fixture('embedded-assets-sample.pptx'), false);
		const handler = new PptxHandler();
		const open = () =>
			handler.load(
				bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
			);
		const { slides } = await open();
		const data = slides
			.flatMap((slide) => slide.elements)
			.find(
				(element): element is ChartPptxElement =>
					element.type === 'chart' && element.chartData?.chartPartPath === 'ppt/charts/whole.xml',
			)!.chartData!;
		expect(data.categories).toEqual([]);
		expect(data.series[0]!.name).toBe('Labels');
		handler.dispose();
	});
});
