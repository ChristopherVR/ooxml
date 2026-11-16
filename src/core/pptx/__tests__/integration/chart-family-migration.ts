import JSZip from 'jszip';
import { writeFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../core/PptxHandler';
import { planWithLeftovers } from '../../core/utils/chart-neutral-save-plan';
import type { ChartPptxElement, PptxChartData, PptxElement } from '../../index';
import { canonicalXmlLines, lineDiff } from './chart-save-parity-canon';
import type { SaveDeck } from './chart-save-parity';

/**
 * Old-versus-new part diff for one PowerPoint chart source adapter moving onto the neutral writer
 * (docs/agnostic-core-plan.md, chart step C3): the same model edit is saved by `PptxHandler.save`
 * twice, once through the neutral plan (the default) and once with the load-time baseline removed,
 * which is what sends a part to the object-tree writers. Both parts are reloaded and compared.
 */

export interface FamilyEdit {
	name: string;
	/** Whether the edit applies to the chart as loaded. */
	applies: (chart: PptxChartData) => boolean;
	edit: (chart: PptxChartData) => void;
	/** The model field the edit sets (compared after reload). */
	read: (chart: PptxChartData) => unknown;
}

type WithRuntime = { runtime: { chartDataBaselines: Map<string, string> } };
const baselines = (handler: PptxHandler) =>
	(handler as unknown as WithRuntime).runtime.chartDataBaselines;

const toBuffer = (bytes: Uint8Array): ArrayBuffer => new Uint8Array(bytes).buffer;

async function loadCharts(bytes: ArrayBuffer) {
	const handler = new PptxHandler();
	const data = await handler.load(bytes);
	const charts = new Map<string, { element: ChartPptxElement; slide: { isDirty?: boolean } }>();
	for (const slide of data.slides) {
		const visit = (list: PptxElement[] | undefined) => {
			for (const element of list ?? []) {
				const chart = element as ChartPptxElement;
				if (element.type === 'chart' && chart.chartData?.chartPartPath)
					charts.set(chart.chartData.chartPartPath, { element: chart, slide });
				visit((element as { children?: PptxElement[] }).children);
			}
		};
		visit(slide.elements);
	}
	return { handler, data, charts };
}

export interface FamilySave {
	part: string;
	xml: string;
	chart: PptxChartData;
	/** `edit.read` of the edited model before the save (what the reload should hold). */
	expected: unknown;
	/** Fields of the edit no neutral family claimed (empty when the neutral plan holds it). */
	unclaimed: string[];
}

/** Saves `edit` on every applicable chart of `deck`, through the neutral plan or the object tree. */
export async function saveFamilyEdit(
	deck: SaveDeck,
	edit: FamilyEdit,
	path: 'neutral' | 'legacy',
): Promise<FamilySave[]> {
	const loaded = await loadCharts(toBuffer(deck.bytes));
	const parts: string[] = [];
	const expected = new Map<string, unknown>();
	const unclaimed = new Map<string, string[]>();
	for (const [part, target] of loaded.charts) {
		const chart = target.element.chartData!;
		if (!/<c:chartSpace[\s>]/u.test(deck.parts.find((p) => p.part === part)?.xml ?? '')) continue;
		if (!edit.applies(chart)) continue;
		const before = structuredClone(chart);
		edit.edit(chart);
		unclaimed.set(part, planWithLeftovers(before, chart).unclaimed);
		expected.set(part, structuredClone(edit.read(chart)));
		target.slide.isDirty = true;
		parts.push(part);
		// Without its load-time baseline a part has no neutral plan, as for an unclaimed edit.
		if (path === 'legacy') baselines(loaded.handler).delete(part);
	}
	if (parts.length === 0) return [];
	const saved = await loaded.handler.save(loaded.data.slides);
	const zip = await JSZip.loadAsync(saved);
	const reloaded = await loadCharts(toBuffer(saved));
	return Promise.all(
		parts.map(async (part) => ({
			part,
			xml: await zip.file(part)!.async('string'),
			chart: reloaded.charts.get(part)!.element.chartData!,
			expected: expected.get(part),
			unclaimed: unclaimed.get(part) ?? [],
		})),
	);
}

/** Canonical XML lines only one of the two saves wrote (`-` legacy, `+` neutral). */
export const familyXmlDifferences = (legacy: string, neutral: string): string[] =>
	lineDiff(canonicalXmlLines(legacy), canonicalXmlLines(neutral));

/** What one family edit did over the corpus. */
export interface FamilyTally {
	/** Parts the edit applied to. */
	parts: number;
	/** Parts whose object-tree save reloads with the edited value. */
	legacyKept: number;
	/** Parts both paths saved to the same canonical XML. */
	xml: number;
}

const emptyTally = (): FamilyTally => ({ parts: 0, legacyKept: 0, xml: 0 });

/**
 * Saves every edit of a family both ways over `decks`: each neutral save must be claimed by the
 * plan and reload with the edited value; whether the object-tree save kept the edit and whether
 * both wrote the same canonical XML is tallied against `expected`, which records what the
 * object-tree writers did (or dropped) for each edit.
 */
export function describeFamilyMigration(
	decks: readonly SaveDeck[],
	edits: readonly FamilyEdit[],
	expected: Record<string, FamilyTally>,
): void {
	const tallies: Record<string, FamilyTally> = {};
	describe.each(edits)('$name', (edit) => {
		it.each(decks)('$name', async (deck) => {
			const [neutral, legacy] = await Promise.all([
				saveFamilyEdit(deck, edit, 'neutral'),
				saveFamilyEdit(deck, edit, 'legacy'),
			]);
			expect(neutral.length).toBe(legacy.length);
			const tally = (tallies[edit.name] ??= emptyTally());
			neutral.forEach((saved, i) => {
				const old = legacy[i]!;
				expect(saved.part).toBe(old.part);
				expect(saved.unclaimed).toEqual([]);
				expect(edit.read(saved.chart)).toEqual(saved.expected);
				tally.parts++;
				if (JSON.stringify(edit.read(old.chart)) === JSON.stringify(old.expected))
					tally.legacyKept++;
				const diff = familyXmlDifferences(old.xml, saved.xml);
				const report = process.env['CHART_FAMILY_REPORT'];
				if (report && process.env['CHART_FAMILY_DETAIL'] === edit.name && diff.length > 0)
					writeFileSync(report, `${[`${deck.name} ${saved.part}`, ...diff].join('\n')}\n`, {
						flag: 'a',
					});
				if (diff.length === 0) tally.xml++;
			});
		});
	});
	it('records the parts compared', () => {
		const all = Object.fromEntries(
			edits.map((edit) => [edit.name, tallies[edit.name] ?? emptyTally()]),
		);
		if (process.env['CHART_FAMILY_REPORT'])
			writeFileSync(process.env['CHART_FAMILY_REPORT'], `${JSON.stringify(all)}\n`, { flag: 'a' });
		expect(all).toEqual(expected);
	});
}
