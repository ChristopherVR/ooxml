import { describe, expect, it } from 'vitest';

import { writeChartSpace } from '../../../chart/index';
import { compareChartSaves } from './chart-save-fixture';
import { classicChartDecks, modelDifferences, xmlDifferences } from './chart-save-parity';
import { SAVE_EDITS } from './chart-save-parity-edits';

/**
 * Chart save harness (docs/agnostic-core-plan.md, chart step C3). For every classic chart part of
 * every committed deck, each edit of `chart-save-parity-edits.ts` is saved through PowerPoint's save
 * and through the neutral reference (`chart/edit` plus `writeChartSpace`); the saved parts are
 * compared as canonical XML and as the reloaded `PptxChartData`. Differences are grouped into
 * categories (element path without values, model path without indexes) per edit.
 */

/**
 * Parts per edit, and how many of them save to the same canonical XML and load back as the same
 * `PptxChartData` both ways. Recorded, so a change on either side shows up here; update the numbers
 * deliberately. On the object-tree path most differences are not the edit itself but what the
 * PowerPoint writers do to the rest of an edited part (see docs/agnostic-core-plan.md, step C3):
 * they renumber series `c:idx`, drop an all-default group `c:dLbls`, pin theme fonts (`+mn-lt`) and
 * resolved run and line colours, rewrite every cache from JavaScript numbers and normalise manual
 * layout numbers. The neutral model keeps the picture options of a 3-D bar series and surface band
 * formats verbatim, so every part is on the neutral path; `bar3d-picture-fill.pptx` writes a series
 * `c:shape` before `c:cat`, which the neutral writer moves to its schema position after `c:val`.
 */
const EXPECTED = {
	'no edit (slide dirty)': { parts: 75, xml: 74, model: 75 },
	'title text': { parts: 75, xml: 75, model: 75 },
	'title hidden': { parts: 75, xml: 75, model: 75 },
	'legend position': { parts: 75, xml: 75, model: 75 },
	// Hiding an already absent legend is a no-op: keep the authored bar3D c:shape ordering.
	'legend hidden': { parts: 75, xml: 74, model: 75 },
	'value axis scaling': { parts: 51, xml: 51, model: 51 },
	'value axis hidden': { parts: 51, xml: 51, model: 51 },
	'value axis number format': { parts: 51, xml: 51, model: 51 },
	'series values': { parts: 75, xml: 75, model: 75 },
	categories: { parts: 71, xml: 71, model: 71 },
	'series name': { parts: 75, xml: 75, model: 75 },
	'data labels': { parts: 72, xml: 72, model: 72 },
	'series colour': { parts: 75, xml: 75, model: 75 },
};

/** An XML difference line reduced to its category: sign, the last three path steps, attribute names. */
function xmlCategory(line: string): string {
	const sign = line.slice(0, 1);
	const body = line.slice(2);
	const space = body.indexOf(' ');
	const pathPart = space < 0 ? body : body.slice(0, space);
	const rest = space < 0 ? '' : body.slice(space + 1);
	const tail = pathPart.split('/').slice(-3).join('/');
	if (pathPart.endsWith('#text')) return `${sign} ${tail}`;
	const names = [...rest.matchAll(/([\w:]+)="/gu)].map((match) => match[1]).join(',');
	return `${sign} ${tail}${names ? ` [${names}]` : ''}`;
}

const modelCategory = (line: string): string =>
	line.slice(0, line.indexOf(':')).replace(/\.\d+(?=\.|$)/gu, '.*');

interface Tally {
	parts: number;
	identicalXml: number;
	identicalModel: number;
	xml: Map<string, number>;
	model: Map<string, number>;
}

const GROUPS = {
	dirty: 'no edit (slide dirty)',
	title: 'title text',
	'title-hidden': 'title hidden',
	legend: 'legend position',
	'legend-hidden': 'legend hidden',
	'axis-display': 'value axis scaling',
	'axis-hidden': 'value axis hidden',
	'axis-format': 'value axis number format',
	values: 'series values',
	categories: 'categories',
	series: 'series name',
	colour: 'series colour',
	labels: 'data labels',
} as const satisfies Record<string, keyof typeof EXPECTED>;
const bump = (map: Map<string, number>, key: string) => map.set(key, (map.get(key) ?? 0) + 1);

export async function describeChartSaveParity(group: keyof typeof GROUPS): Promise<void> {
	const decks = await classicChartDecks();
	const name = GROUPS[group];
	const edits = SAVE_EDITS.filter((edit) => edit.name === name);
	const TALLY = new Map<string, Tally>();
	describe('chart save: PowerPoint writers vs neutral writeChartSpace', () => {
		describe.each(edits)('$name', (edit) => {
			it.each(decks)('$name', async (deck) => {
				const result = await compareChartSaves(deck, edit, writeChartSpace);
				if (!result) return;
				const { parts, fromPptx, fromNeutral } = result;
				const tally = TALLY.get(edit.name) ?? {
					parts: 0,
					identicalXml: 0,
					identicalModel: 0,
					xml: new Map(),
					model: new Map(),
				};
				TALLY.set(edit.name, tally);
				for (const part of parts) {
					const a = fromPptx.get(part)!;
					const b = fromNeutral.get(part)!;
					tally.parts++;
					const xml = xmlDifferences(a.xml, b.xml);
					const model = modelDifferences(a.chart, b.chart);
					if (process.env['CHART_SAVE_DETAIL'] === edit.name && (xml.length || model.length))
						console.log(
							[`${deck.name} ${part}`, ...xml, ...model].map((l) => l.slice(0, 300)).join('\n'),
						);
					if (xml.length === 0) tally.identicalXml++;
					if (model.length === 0) tally.identicalModel++;
					for (const key of new Set(xml.map(xmlCategory))) bump(tally.xml, key);
					for (const key of new Set(model.map(modelCategory))) bump(tally.model, key);
				}
				expect(tally.parts).toBeGreaterThan(0);
			});
		});

		it('reports the difference categories', () => {
			const lines: string[] = [];
			for (const [name, tally] of TALLY) {
				lines.push(
					`${name}: ${String(tally.parts)} parts, xml identical ${String(tally.identicalXml)}, model identical ${String(tally.identicalModel)}`,
				);
				for (const [key, count] of [...tally.xml].sort())
					lines.push(`  xml   ${String(count).padStart(3)} ${key}`);
				for (const [key, count] of [...tally.model].sort())
					lines.push(`  model ${String(count).padStart(3)} ${key}`);
			}
			if (process.env['CHART_SAVE_REPORT']) console.log(lines.join('\n'));
			const totals = Object.fromEntries(
				[...TALLY].map(([name, t]) => [
					name,
					{ parts: t.parts, xml: t.identicalXml, model: t.identicalModel },
				]),
			);
			expect(totals).toEqual({ [name]: EXPECTED[name] });
		});
	});
	if (group === 'dirty') {
		it('assigns each edit and its expected tally to exactly one file', () => {
			const names = Object.values(GROUPS).sort();
			expect(names).toEqual(SAVE_EDITS.map(({ name }) => name).sort());
			expect(names).toEqual(Object.keys(EXPECTED).sort());
		});
	}
}
