import { writeFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

import type { PptxHandler } from '../../core/PptxHandler';
import { planWithLeftovers } from '../../core/utils/chart-neutral-save-plan';
import type { PptxChartData } from '../../index';
import { loadChartDeck, readSavedCharts, type SavedChart } from './chart-save-fixture';
export { xmlDifferences as familyXmlDifferences } from './chart-save-parity-canon';
import { xmlDifferences } from './chart-save-parity-canon';
import { classicChartDecks, type SaveDeck } from './chart-save-parity';

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

const sourceCharts = new WeakMap<SaveDeck, Promise<ReadonlyMap<string, PptxChartData>>>();

/** Cache applicability inputs only. Each writer below still gets its own load and save. */
function chartsIn(deck: SaveDeck): Promise<ReadonlyMap<string, PptxChartData>> {
	let charts = sourceCharts.get(deck);
	if (!charts) {
		charts = loadChartDeck(deck.bytes).then((loaded) => {
			const classicParts = new Set(deck.parts.map(({ part }) => part));
			const out = new Map<string, PptxChartData>();
			for (const [part, { element }] of loaded.charts)
				if (classicParts.has(part)) out.set(part, element.chartData);
			loaded.handler.dispose();
			return out;
		});
		sourceCharts.set(deck, charts);
	}
	return charts;
}

export interface FamilySave extends SavedChart {
	part: string;
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
	const targets = [...(await chartsIn(deck))].filter(([, chart]) => edit.applies(chart));
	if (targets.length === 0) return [];
	const loaded = await loadChartDeck(deck.bytes);
	try {
		const parts: string[] = [];
		const expected = new Map<string, unknown>();
		const unclaimed = new Map<string, string[]>();
		for (const [part] of targets) {
			const target = loaded.charts.get(part)!;
			const chart = target.element.chartData;
			const before = structuredClone(chart);
			edit.edit(chart);
			unclaimed.set(part, planWithLeftovers(before, chart).unclaimed);
			expected.set(part, structuredClone(edit.read(chart)));
			target.slide.isDirty = true;
			parts.push(part);
			// Without its load-time baseline a part has no neutral plan, as for an unclaimed edit.
			if (path === 'legacy') baselines(loaded.handler).delete(part);
		}
		const saved = await loaded.handler.save(loaded.data.slides);
		const reloaded = await readSavedCharts(saved, parts);
		return parts.map((part) => ({
			part,
			...reloaded.get(part)!,
			expected: expected.get(part),
			unclaimed: unclaimed.get(part) ?? [],
		}));
	} finally {
		loaded.handler.dispose();
	}
}

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

/** Register full corpus tallies for the edits assigned to one file. */
export async function describeFamilyMigrationEdits(
	edits: readonly FamilyEdit[],
	expected: Record<string, FamilyTally>,
	names: readonly string[],
): Promise<void> {
	describeFamilyMigration(
		await classicChartDecks(),
		edits.filter(({ name }) => names.includes(name)),
		Object.fromEntries(names.map((name) => [name, expected[name]])),
	);
}

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
				const diff = xmlDifferences(old.xml, saved.xml);
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
