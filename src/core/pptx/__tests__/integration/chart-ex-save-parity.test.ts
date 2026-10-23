import { describe, expect, it } from 'vitest';

import { modelDifferences, xmlDifferences } from './chart-save-parity';
import {
	CHART_EX_SAVE_EDITS,
	chartExDecks,
	chartExSavedByNeutral,
	chartExSavedByPptx,
	chartExTargets,
} from './chart-ex-save-parity';

/**
 * ChartEx save harness (docs/agnostic-core-plan.md, chart step C5). For every committed chartex
 * part (15, the Word and Excel parts swapped into the gallery deck), each edit is saved through
 * PowerPoint's save and through the neutral reference (`chart/edit` chartex operations plus
 * `writeChartExSpace`); the saved parts are compared as canonical XML and as the reloaded
 * `PptxChartData`. The recorded totals pin the result; update them deliberately.
 */

const DECKS = await chartExDecks();

/**
 * Parts per edit, and how many save to the same canonical XML and load back as the same
 * `PptxChartData` both ways. Before the claimed families went through `writeChartExSpace`
 * (`utils/chart-ex-neutral-save.ts`), the object-tree update matched on no part's XML for the
 * title, legend-hidden, series-name and data-label edits (3 to 5 of 15 for the others): on every
 * edit it wrote `formatCode="General"` on every numeric level, turned a `cx:txData` title into
 * rich text and gave a chart without categories an empty category dimension. The one difference
 * left is the region map's categories, which the plan does not claim (PowerPoint writes them with
 * their source indexes), so the object-tree update still saves that edit.
 */
const EXPECTED: Record<string, { parts: number; xml: number; model: number }> = {
	'no edit (slide dirty)': { parts: 15, xml: 15, model: 15 },
	'title text': { parts: 15, xml: 15, model: 15 },
	'title hidden': { parts: 15, xml: 15, model: 15 },
	'legend position': { parts: 15, xml: 15, model: 15 },
	'legend hidden': { parts: 15, xml: 15, model: 15 },
	'series name': { parts: 15, xml: 15, model: 15 },
	'series values': { parts: 15, xml: 15, model: 15 },
	categories: { parts: 6, xml: 5, model: 6 },
	'data labels': { parts: 15, xml: 15, model: 15 },
};

const TALLY = new Map<string, { parts: number; xml: number; model: number }>();

describe('chartex save: PowerPoint writers vs neutral writeChartExSpace', () => {
	describe.each(CHART_EX_SAVE_EDITS)('$name', (edit) => {
		it.each(DECKS)('$name', async (deck) => {
			const targets = await chartExTargets(deck, edit);
			if (targets.length === 0) return;
			const [fromPptx, fromNeutral] = await Promise.all([
				chartExSavedByPptx(deck, edit, targets),
				chartExSavedByNeutral(deck, edit, targets),
			]);
			const tally = TALLY.get(edit.name) ?? { parts: 0, xml: 0, model: 0 };
			TALLY.set(edit.name, tally);
			for (const target of targets) {
				const a = fromPptx.get(target.part)!;
				const b = fromNeutral.get(target.part)!;
				tally.parts++;
				const xml = xmlDifferences(a.xml, b.xml);
				const model = modelDifferences(a.chart, b.chart);
				if (process.env['CHART_SAVE_DETAIL'] === edit.name && (xml.length || model.length))
					console.log(
						[`${deck.name} ${target.part}`, ...xml, ...model]
							.map((line) => line.slice(0, 300))
							.join('\n'),
					);
				if (xml.length === 0) tally.xml++;
				if (model.length === 0) tally.model++;
			}
		});
	});

	it('records the totals', () => {
		const totals = Object.fromEntries(TALLY);
		if (process.env['CHART_SAVE_REPORT']) console.log(JSON.stringify(totals, null, 1));
		expect(totals).toEqual(EXPECTED);
	});
});
