// The xlsx chart save parity harness (chart step C2). Before charts were edited over `ChartSpace`,
// `updateChart` patched kept parts through the DOM (`write/chart-patch.ts`, `write/chart-colors.ts`).
// `__fixtures__/chart-save-parity.json` holds, for every chart of the corpus and every edit kind
// (`chart-parity-cases.ts`), a fingerprint of what that path saved: `-` when the edit did not
// apply, `=` when the part was saved byte for byte, else a hash of the saved part's model without
// its `sourceXml` copies (which hold the same XML with namespace declarations placed differently).
// The operations path must save the same model, except where a difference is listed below with
// its reason and checked for what the edit asked.
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { expect, it } from 'vitest';
import type { ChartSpace } from '../../../chart/model';
import { parseChartSpace } from '../../../chart/parse-space';
import { chartSummaryFromSpace } from '../../../chart/render/summary-from-space';
import { PARITY_EDITS, parityCharts, runParityEdit, type ParityChart } from './chart-parity-cases';

const FIXTURE = path.join(import.meta.dirname, '../../__fixtures__/chart-save-parity.json');

const stripped = (space: ChartSpace): unknown =>
	JSON.parse(
		JSON.stringify(space, (key, value) => {
			if (key === 'sourceXml') return undefined;
			// The frozen DOM-path hashes predate these authored outline fields. Keep their
			// original comparison scope; DrawingML's real-fixture round trips separately
			// verify the new fields and their preservation through chart edit/save.
			if (key === 'line' && value && typeof value === 'object') {
				const {
					compound: _compound,
					alignment: _alignment,
					customDash: _dash,
					join: _join,
					headEnd: _head,
					tailEnd: _tail,
					...legacy
				} = value;
				return legacy;
			}
			return value;
		}),
	);

/** The hash of a part's model without its `sourceXml` copies. */
function modelHash(xml: string): string {
	const model = stable(stripped(parseChartSpace(xml).chartSpace));
	return createHash('sha256').update(JSON.stringify(model)).digest('hex').slice(0, 16);
}

/** The fingerprint of a saved part (see the file comment). */
export function chartPartFingerprint(source: string, saved: string): string {
	return saved === source ? '=' : modelHash(saved);
}

function stable(value: unknown): unknown {
	if (Array.isArray(value)) return value.map(stable);
	if (value && typeof value === 'object')
		return Object.fromEntries(
			Object.keys(value)
				.sort()
				.map((key) => [key, stable((value as Record<string, unknown>)[key])]),
		);
	return value;
}

const formulas = (space: ChartSpace) =>
	space.plotArea.groups.flatMap((group) =>
		group.series.flatMap((series) =>
			[
				series.tx?.reference,
				series.categories,
				series.values,
				series.xValues,
				series.yValues,
				series.bubbleSizes,
			].flatMap((source) => (source?.formula ? [source.formula] : [])),
		),
	);

type Check = (saved: ChartSpace, source: ChartSpace) => void;

const summaryOf = (space: ChartSpace) => chartSummaryFromSpace(space);

/**
 * Edit kinds the operations save differently, by design. The DOM path regenerated the part from
 * the summary for a type, grouping or series-count change (losing every unmodelled detail and
 * turning a stock chart into columns); the operations keep the part. It also left `c:delete` alone
 * on axes without direct formatting, and rewrote the references of the first chart group only.
 */
export const KNOWN_DIFFERENCES: Record<string, { reason: string; check: Check }> = {
	typeLine: {
		reason: 'setChartType keeps the formatting; the DOM path regenerated the part',
		check: (saved, source) => {
			const before = summaryOf(source);
			expect(summaryOf(saved).chartType).toBe(before.chartType === 'line' ? 'column' : 'line');
			// A combination chart becomes one group holding every series.
			const values = (space: ChartSpace) =>
				space.plotArea.groups.flatMap((group) =>
					group.series.map((series) => (series.values ?? series.yValues)?.formula),
				);
			expect(values(saved)).toEqual(values(source));
		},
	},
	groupingStacked: {
		reason: 'setChartGrouping keeps the part; the DOM path regenerated it',
		check: (saved) => expect(summaryOf(saved).grouping).toBe('stacked'),
	},
	addSeries: {
		reason: 'addChartSeries keeps the part; the DOM path regenerated it',
		check: (saved, source) => {
			const series = summaryOf(saved).series;
			expect(series).toHaveLength(summaryOf(source).series.length + 1);
			expect(series.at(-1)?.name).toBe('Added series');
		},
	},
	removeSeries: {
		reason: 'removeChartSeries keeps the part; the DOM path regenerated it',
		check: (saved, source) => {
			const before = summaryOf(source).series;
			expect(summaryOf(saved).series.map((s) => s.valuesRef)).toEqual(
				before.slice(1).map((s) => s.valuesRef),
			);
		},
	},
	valueAxisHidden: {
		reason:
			'the axis edit is written as c:delete; the DOM path ignored axes without direct formatting',
		check: (saved) => {
			for (const axis of saved.plotArea.axes.filter((item) => item.kind === 'val'))
				expect(axis.deleted).toBe(true);
		},
	},
	insertRows: {
		reason: 'every group and bubble-size reference moves; the DOM path moved the first group only',
		check: (saved, source) => {
			const shifted = formulas(source).map((formula) =>
				formula.replace(/\$(\d+)/g, (_, row: string) => `$${Number(row) + 1}`),
			);
			expect(formulas(saved)).toEqual(shifted);
		},
	},
};

/**
 * Single cases that differ: a point whose marker already shows the colour the edit gives it is not
 * written again as a `c:dPt/c:spPr` fill (the painted chart is the same).
 */
const greenPoint: Check = (saved) =>
	expect(summaryOf(saved).series[0]?.pointColors?.[1]).toEqual({
		kind: 'srgb',
		value: '00FF00',
		transforms: [],
	});

export const CASE_DIFFERENCES: Record<string, Check> = {
	'excel-chart-palette-edits.json#26 :: pointColor': greenPoint,
	'excel-chart-palette-edits.json#27 :: pointColor': greenPoint,
};

/** The recorded fingerprints: chart id to one entry per edit kind, in `PARITY_EDITS` order. */
export function recordedFingerprints(): Record<string, string[]> {
	const data = JSON.parse(readFileSync(FIXTURE, 'utf8')) as Record<string, string>;
	return Object.fromEntries(Object.entries(data).map(([id, row]) => [id, row.split(' ')]));
}

/** One test per chart whose id `include` accepts: every edit kind saved and compared. */
export function describeChartSaveParity(include: (id: string) => boolean): void {
	const recorded = recordedFingerprints();
	const kinds = Object.keys(PARITY_EDITS);
	const ids = Object.keys(recorded).filter(include);
	let charts: Promise<ParityChart[]> | undefined;
	it.each(ids)('saves %s like the DOM patch path did', async (id) => {
		const chart = (await (charts ??= parityCharts())).find((item) => item.id === id);
		expect(chart, `chart ${id} is in the corpus`).toBeDefined();
		const row = recorded[id]!;
		for (const [i, kind] of kinds.entries()) {
			const result = await runParityEdit(chart!, PARITY_EDITS[kind]!);
			const key = `${id} :: ${kind}`;
			const expected = row[i];
			if (expected === '-') {
				expect(result, key).toBeUndefined();
				continue;
			}
			expect(result, key).toBeDefined();
			const { source, saved } = result!;
			// The same model, or the source itself where the DOM path saved it byte for byte.
			const same = expected === '=' ? saved === source : modelHash(saved) === expected;
			if (same) continue;
			const difference = KNOWN_DIFFERENCES[kind]?.check ?? CASE_DIFFERENCES[key];
			if (difference) {
				difference(parseChartSpace(saved).chartSpace, parseChartSpace(source).chartSpace);
				continue;
			}
			expect(chartPartFingerprint(source, saved), key).toBe(expected);
		}
	});
}
