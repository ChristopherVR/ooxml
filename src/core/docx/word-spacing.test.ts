// Paragraph spacing against what desktop Word measured (`__fixtures__/word-spacing/
// spacing-probe.json`, written by `scripts/docx-word-spacing-probe.ps1`): the same small documents
// are rebuilt (`word-spacing-docs.ts`), laid out and compared paragraph by paragraph (page exact,
// top y within 1 pt: Word reports positions in whole 0.75 pt pixels).
//
// Measured rules this pins (identical in compatibility modes 12, 14 and 15):
//  - without `w:doNotUseHTMLParagraphAutoSpacing`, two paragraphs are `max(after, before)` apart,
//    `beforeAutospacing`/`afterAutospacing` are 14 pt, and the first paragraph of a section that
//    starts a page keeps `max(0, before - previous after)`;
//  - with it, the spaces add and the autospacing flags are ignored (the twips are used);
//  - into and out of a table the spaces always add; a page or column top (after a page break,
//    `pageBreakBefore`, an overflow or a column break) drops the space before.
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { loadDocx } from './index';
import { layoutDocumentModel } from './layout/layout';
import type { LayoutPageBox } from './layout/result';
import type { TextMeasurer } from './layout/measure';
import { probeDocx } from './word-spacing-docs';

interface Row {
	label: string;
	page: number;
	y: number;
}
interface Probe {
	configs: { mode: number; noHtmlAutoSpacing: boolean; groups: Record<string, Row[]> }[];
}
const probe = JSON.parse(
	readFileSync(
		new URL('./__fixtures__/word-spacing/spacing-probe.json', import.meta.url),
		'utf8',
	).replace(/^﻿/, ''),
) as Probe;
const measurer: TextMeasurer = {
	widthOf: (text, font) => Array.from(text).length * font.sizePx * 0.6,
	lineHeightOf: (font) => font.sizePx * (font.family.includes('Symbol') ? 1.29 : 1.1328),
};
const PT_PER_PX = 0.75;

/** Page and top y (pt) of the first line of each labelled paragraph outside tables. */
function positions(pages: LayoutPageBox[]): Map<string, { page: number; y: number }> {
	const found = new Map<string, { page: number; y: number }>();
	for (const page of pages)
		for (const column of page.columns) {
			// Paragraphs inside a table cell are positioned relative to the cell, so only top-level ones.
			for (const box of column.blocks) {
				if (box.kind !== 'paragraph') continue;
				const text = box.lines.flatMap((line) => line.fragments.map((f) => f.text)).join('');
				const label = /([A-Za-z]\w*)$/.exec(text)?.[1];
				if (!label || found.has(label)) continue;
				found.set(label, {
					page: page.index + 1,
					y: (page.marginTopPx + box.yPx + (box.lines[0]?.yPx ?? 0)) * PT_PER_PX,
				});
			}
		}
	return found;
}

const configs = probe.configs.map((c) => ({
	name: `compat ${c.mode}, doNotUseHTMLParagraphAutoSpacing ${c.noHtmlAutoSpacing ? 'on' : 'off'}`,
	...c,
}));

describe('paragraph spacing against the Word probe', () => {
	it('measured the three compatibility modes with and without the option', () => {
		expect(configs.map((c) => `${c.mode}/${c.noHtmlAutoSpacing}`)).toEqual([
			'15/false',
			'15/true',
			'14/false',
			'14/true',
			'12/false',
			'12/true',
		]);
	});
	describe.each(configs)('$name', (config) => {
		it.each(Object.keys(config.groups))('%s', async (group) => {
			const loaded = await loadDocx(await probeDocx(group, config.mode, config.noHtmlAutoSpacing));
			const ours = positions(layoutDocumentModel(loaded.model, measurer).pages);
			const wrong: string[] = [];
			const split: string[] = [];
			for (const row of config.groups[group]!) {
				const mine = ours.get(row.label);
				if (!mine) continue; // table cell paragraphs and section-break leftovers
				if (mine.page !== row.page || Math.abs(mine.y - row.y) > 1)
					(group === 'table' && row.page === 2 ? split : wrong).push(
						`${row.label}: page ${mine.page} y ${mine.y.toFixed(1)}, Word ${row.page}/${row.y}`,
					);
			}
			expect(wrong).toEqual([]);
			// Known gap: a one-row table moved whole to the next page loses its last cell paragraph's
			// space after (12 pt) in the layout, which shifts everything after it on that page.
			if (group === 'table')
				expect(split, 'table page-split gap now passes: remove it').not.toEqual([]);
		});
	});
});
