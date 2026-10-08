import JSZip from 'jszip';
import { readdirSync, readFileSync } from 'node:fs';
import path from 'node:path';

import { describe, expect, it } from 'vitest';

import { parseChartSpace } from '../../../chart/index';
import { PptxHandler } from '../../index';
import type { PptxChartData, PptxElement } from '../../index';
import { diffShapes } from './chart-neutral-parity';
import { projectNeutralChart } from './chart-neutral-parity-neutral';
import { projectPptxChart } from './chart-neutral-parity-pptx';

/**
 * Chart parity harness (docs/agnostic-core-plan.md, step 5 and "Risks"): before pptx adopts the
 * neutral chart model, every chart part of every committed deck is read by BOTH parsers - the deck
 * through `PptxHandler` (the `fast-xml-parser` object tree, `PptxChartData`) and the raw part
 * through `parseChartSpace` (the shared DOM, `ChartSpace`) - and both are projected onto one
 * comparison shape (`chart-neutral-parity.ts`) that must be equal.
 *
 * pptx now builds its classic chart model from `parseChartSpace` (`utils/chart-from-neutral*.ts`),
 * so this harness pins the adapter's normalisations against an independent projection of the
 * neutral model.
 *
 * Where the two models legitimately differ, the difference is an explicit expectation below, not a
 * loose tolerance: each one asserts the difference is still there, so a fix on either side turns the
 * suite red and tells you to delete the expectation.
 */

const ROOT = path.resolve(__dirname, '../fixtures');
const CHART_PART = /^ppt\/charts\/chart[^/]*\.xml$/u;
/** The ChartEx (`cx:`) layouts pptx models and the neutral parser only reports. */
const CHART_EX_TYPES = new Set([
	'waterfall',
	'funnel',
	'treemap',
	'sunburst',
	'boxWhisker',
	'histogram',
	'regionMap',
]);

function walk(dir: string): string[] {
	return readdirSync(dir, { withFileTypes: true })
		.flatMap((entry) => {
			const full = path.join(dir, entry.name);
			if (entry.isDirectory()) return walk(full);
			return entry.name.endsWith('.pptx') ? [full] : [];
		})
		.sort();
}

interface ChartDeck {
	name: string;
	bytes: Buffer;
	zip: JSZip;
	parts: string[];
}

/** Every committed deck that holds at least one chart part (encrypted and legacy decks are skipped). */
async function chartDecks(): Promise<ChartDeck[]> {
	const decks: ChartDeck[] = [];
	for (const file of walk(ROOT)) {
		const bytes = readFileSync(file);
		let zip: JSZip;
		try {
			zip = await JSZip.loadAsync(bytes);
		} catch {
			continue;
		}
		const parts = Object.keys(zip.files)
			.filter((name) => CHART_PART.test(name))
			.sort();
		if (parts.length > 0)
			decks.push({ name: path.relative(ROOT, file).replace(/\\/gu, '/'), bytes, zip, parts });
	}
	return decks;
}

/** The pptx chart model of every chart frame on every slide (groups included), by part name. */
async function pptxChartsByPart(bytes: Buffer): Promise<Map<string, PptxChartData>> {
	const data = await new PptxHandler().load(
		bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
	);
	const byPart = new Map<string, PptxChartData>();
	const visit = (list: PptxElement[] | undefined) => {
		for (const element of list ?? []) {
			const chartData = (element as { chartData?: PptxChartData }).chartData;
			if (element.type === 'chart' && chartData?.chartPartPath)
				byPart.set(chartData.chartPartPath, chartData);
			visit((element as { children?: PptxElement[] }).children);
		}
	};
	for (const slide of data.slides) visit(slide.elements);
	return byPart;
}

const DECKS = await chartDecks();
const totals = { parts: 0, classic: 0, chartEx: 0, series: 0, axes: 0 };

describe('chart parity: pptx object-tree parser vs neutral parseChartSpace', () => {
	describe.each(DECKS)('$name', (deck) => {
		it('projects every chart part identically', async () => {
			const byPart = await pptxChartsByPart(deck.bytes);
			for (const part of deck.parts) {
				const pptx = byPart.get(part);
				// Every chart part of the corpus is reached from a slide frame.
				expect(pptx, `${part} is not reached from any slide`).toBeDefined();
				if (!pptx) continue;
				totals.parts++;
				const xml = await deck.zip.file(part)!.async('string');
				const { chartSpace, issues } = parseChartSpace(xml);

				// Expectation (gap, not a bug): a ChartEx part (`cx:chartSpace`, here stored under a
				// chartN.xml name) is modelled by pptx (`parseCxChartSeries`); the neutral parser reports
				// it and returns an empty plot area.
				if (CHART_EX_TYPES.has(pptx.chartType)) {
					totals.chartEx++;
					expect(issues.map((issue) => issue.code)).toContain('CHART_ROOT_UNEXPECTED');
					expect(chartSpace.plotArea.groups).toEqual([]);
					continue;
				}
				totals.classic++;

				const neutral = projectNeutralChart(chartSpace);
				const legacy = projectPptxChart(pptx);
				totals.series += legacy.series.length;
				totals.axes += legacy.axes.length;

				// Expectation (pptx normalisation, `getChartDataForGraphicFrame`): when the part caches
				// no categories, pptx takes them from the embedded workbook's first column, which the
				// neutral parser (a part parser) never opens.
				const workbookCategories = pptx.embeddedWorkbookData?.categories ?? [];
				if (neutral.categories.length === 0 && workbookCategories.length > 0) {
					expect(legacy.categories).toEqual(workbookCategories);
					neutral.categories = workbookCategories;
				}

				expect(diffShapes(neutral, legacy), `${deck.name} ${part}`).toEqual([]);
			}
		});
	});

	// Guards the corpus against silently shrinking: a deck or chart part that disappears, or a
	// classic part that turns into an unreadable one, changes these totals. Update them deliberately
	// when a chart fixture is added.
	it('covered every chart part in the corpus', () => {
		expect(DECKS).toHaveLength(23);
		expect(totals).toEqual({ parts: 82, classic: 75, chartEx: 7, series: 148, axes: 110 });
	});
});
