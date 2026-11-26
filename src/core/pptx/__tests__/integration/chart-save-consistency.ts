// Every chart save writes ranges and caches that agree with the embedded workbook, whatever the
// undo history kept and however much of it the saved-state budget and garbage collection left.
// CI runs a few cases that reach every fallback; `CHART_SAVE_SWEEP=full` runs every committed
// chart deck under every history copy, budget and collection timing, and
// `CHART_SAVE_SWEEP_OUT=<folder>` writes each saved deck there, for opening in PowerPoint, and a
// summary per group with every case, its saves, issues and time. Use the `chart-save-consistency`
// filename prefix to run all groups.
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
import { requireFixture } from '../require-fixture';

const FULL = process.env['CHART_SAVE_SWEEP'] === 'full';
const OUT = process.env['CHART_SAVE_SWEEP_OUT'];

const fixtures = new Map<string, Uint8Array>();

function fixture(name: string): Uint8Array {
	let bytes = fixtures.get(name);
	if (!bytes) {
		const url = new URL(`../fixtures/${name}`, import.meta.url);
		bytes = new Uint8Array(readFileSync(requireFixture(fileURLToPath(url))));
		fixtures.set(name, bytes);
	}
	return bytes;
}

const wholeColumnDecks = new Map<boolean, ReturnType<typeof wholeColumnChartDeck>>();

export function wholeColumnFixture(categories: boolean) {
	let deck = wholeColumnDecks.get(categories);
	if (!deck) {
		deck = wholeColumnChartDeck(fixture('embedded-assets-sample.pptx'), categories);
		wholeColumnDecks.set(categories, deck);
	}
	return deck;
}

type SweepGroup = 'chart1' | 'chart3' | 'whole-column';
type SweepSegment = 'history' | 'json-history' | 'replay';
const FULL_EXTRA_DECKS = [
	'chart-ext-datalabels-range',
	'chart-ext-filtered-titles',
	'chart-filtered-series',
	'chart-label-callouts',
];

async function decks(group: SweepGroup) {
	if (group === 'whole-column') {
		const columns = FULL ? [true, false] : [true];
		return Promise.all(
			columns.map(async (categories) => {
				const { bytes, chartPartPath } = await wholeColumnFixture(categories);
				return {
					deck: categories ? 'whole-column' : 'whole-column-text',
					bytes,
					charts: [chartPartPath],
				};
			}),
		);
	}

	const names = ['chart-data-fidelity', ...(FULL && group === 'chart1' ? FULL_EXTRA_DECKS : [])];
	const list: { deck: string; bytes: Uint8Array; charts: string[] }[] = [];
	for (const deck of names) {
		const bytes = fixture(`e2e/${deck}.pptx`);
		if (!FULL) {
			list.push({ deck, bytes, charts: [`ppt/charts/${group}.xml`] });
			continue;
		}
		const charts = await sweepableCharts(bytes);
		const selected =
			deck === 'chart-data-fidelity'
				? charts.filter((chart) => (chart === 'ppt/charts/chart3.xml') === (group === 'chart3'))
				: charts;
		list.push({ deck, bytes, charts: selected });
	}
	return list;
}

type Variant = Pick<SweepCase, 'copy' | 'budget' | 'gc'>;

// CI: kept states replayed exactly, the JSON history at the default budget that lost its states
// to collection, and a budget that keeps none, on PowerPoint column and bubble charts and a
// whole-column chart.
const CI: Record<string, Variant[]> = {
	'chart-data-fidelity': [
		{ copy: 'clone', budget: undefined, gc: true },
		{ copy: 'json', budget: undefined, gc: true },
		{ copy: 'clone', budget: 0, gc: false },
	],
	'whole-column': [
		{ copy: 'json', budget: 0, gc: true },
		{ copy: 'clone', budget: undefined, gc: false },
	],
};

function everyVariant(): Variant[] {
	const all: Variant[] = [];
	for (const copy of ['clone', 'structured', 'json'] as HistoryCopy[])
		for (const budget of [undefined, 0, 2048])
			for (const gc of [true, false]) all.push({ copy, budget, gc });
	return all;
}

async function cases(group: SweepGroup, segment: SweepSegment): Promise<SweepCase[]> {
	const all: SweepCase[] = [];
	for (const { deck, bytes, charts } of await decks(group)) {
		const variants = FULL ? everyVariant() : (CI[deck] ?? []);
		const selected = variants.filter((variant) => segmentFor(group, variant) === segment);
		for (const chartPartPath of charts)
			for (const variant of selected) all.push({ deck, bytes, chartPartPath, ...variant });
	}
	return all;
}

function segmentFor(group: SweepGroup, variant: Variant): SweepSegment {
	if (variant.budget === 0) return 'replay';
	if (group !== 'whole-column' && variant.copy === 'json') return 'json-history';
	return 'history';
}

export function describeChartSaveConsistency(group: SweepGroup, segment: SweepSegment): void {
	describe(`chart save consistency: ${group} ${segment}`, () => {
		it(
			'saves ranges, caches and workbook cells that agree after every edit, undo and redo',
			async () => {
				if (OUT) mkdirSync(OUT, { recursive: true });
				const problems: string[] = [];
				const report: string[] = [];
				let saves = 0;
				for (const c of await cases(group, segment)) {
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
							writeFileSync(`${OUT}/${name.replace(/[/:]/gu, '_')}-${label}.pptx`, bytes);
						}
					}
				}
				report.push(`chart save sweep: ${saves} saves, ${problems.length} problems`, ...problems);
				if (OUT) writeFileSync(`${OUT}/summary-${group}-${segment}.txt`, report.join('\n') + '\n');
				expect(saves).toBeGreaterThan(0);
				expect(problems).toEqual([]);
			},
			FULL ? 3_600_000 : 300_000,
		);
	});
}
