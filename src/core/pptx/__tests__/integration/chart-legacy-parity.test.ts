import JSZip from 'jszip';
import { readdirSync, readFileSync } from 'node:fs';
import path from 'node:path';

import { afterEach, describe, expect, it } from 'vitest';

import { setLegacyChartParsing } from '../../core/core/runtime/legacy-chart-parsing';
import { PptxHandler } from '../../index';
import type { PptxChartData, PptxElement } from '../../index';

/**
 * Old vs new chart parser (docs/agnostic-core-plan.md, steps 5 and 6): pptx now builds
 * `PptxChartData` from the neutral `parseChartSpace` model through the adapter
 * (`utils/chart-from-neutral*.ts`). For one wave the previous object-tree parser is kept as
 * `legacy-chart-parsing.ts`; this test loads every committed deck with a chart part through both
 * and asserts the FULL chart model of every chart frame is deeply equal (not a projection).
 */

const ROOT = path.resolve(__dirname, '../fixtures');
const CHART_PART = /^ppt\/charts\/chart[^/]*\.xml$/u;

function walk(dir: string): string[] {
	return readdirSync(dir, { withFileTypes: true })
		.flatMap((entry) => {
			const full = path.join(dir, entry.name);
			if (entry.isDirectory()) return walk(full);
			return entry.name.endsWith('.pptx') ? [full] : [];
		})
		.sort();
}

async function chartDecks(): Promise<{ name: string; bytes: Buffer }[]> {
	const decks: { name: string; bytes: Buffer }[] = [];
	for (const file of walk(ROOT)) {
		const bytes = readFileSync(file);
		let zip: JSZip;
		try {
			zip = await JSZip.loadAsync(bytes);
		} catch {
			continue;
		}
		if (Object.keys(zip.files).some((name) => CHART_PART.test(name)))
			decks.push({ name: path.relative(ROOT, file).split(path.sep).join('/'), bytes });
	}
	return decks;
}

/** Every chart frame's model, keyed by slide and frame position (groups included). */
async function chartsOf(bytes: Buffer, legacy: boolean): Promise<Map<string, PptxChartData>> {
	setLegacyChartParsing(legacy);
	try {
		const data = await new PptxHandler().load(
			bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
		);
		const charts = new Map<string, PptxChartData>();
		const visit = (list: PptxElement[] | undefined, at: string) =>
			(list ?? []).forEach((element, index) => {
				const chartData = (element as { chartData?: PptxChartData }).chartData;
				if (element.type === 'chart' && chartData) charts.set(`${at}/${index}`, chartData);
				visit((element as { children?: PptxElement[] }).children, `${at}/${index}`);
			});
		data.slides.forEach((slide, index) => visit(slide.elements, `slide${index + 1}`));
		return charts;
	} finally {
		setLegacyChartParsing(false);
	}
}

const DECKS = await chartDecks();
const totals = { decks: 0, charts: 0 };

afterEach(() => setLegacyChartParsing(false));

describe('chart parser: legacy object tree vs neutral adapter', () => {
	describe.each(DECKS)('$name', (deck) => {
		it('builds the same full chart model for every chart frame', async () => {
			const legacy = await chartsOf(deck.bytes, true);
			const current = await chartsOf(deck.bytes, false);
			expect([...current.keys()]).toEqual([...legacy.keys()]);
			for (const [key, chart] of legacy) {
				expect(current.get(key), `${deck.name} ${key} ${chart.chartPartPath}`).toStrictEqual(chart);
			}
			totals.decks++;
			totals.charts += legacy.size;
		});
	});

	it('covered every chart frame in the corpus', () => {
		expect(totals.decks).toBe(DECKS.length);
		expect(totals).toEqual({ decks: 22, charts: 81 });
	});
});
