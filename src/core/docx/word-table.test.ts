// Table row sizing against what desktop Word measured (`__fixtures__/word-table/table-probe.json`,
// written by `scripts/docx-word-table-probe.ps1`): the same small documents are rebuilt
// (`word-table-docs.ts`), laid out and compared paragraph by paragraph (page exact, top y within
// 1 pt: Word reports positions in whole 0.75 pt pixels).
//
// Measured rules this pins:
//  - the text of a vertically merged cell that is taller than the rows it spans grows the LAST
//    spanned row (the rows above keep the height of their own cells);
//  - a row is as tall as its tallest non-merged cell plus the width of its (shared) border, and the
//    first row also starts one border width lower.
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { loadDocx } from './index';
import { layoutDocumentModel } from './layout/layout';
import type { LayoutBlockBox, LayoutPageBox } from './layout/result';
import { wordProbeMeasurer } from './word-measurer';
import { tableBodies, tableDocx } from './word-table-docs';

interface Row {
	label: string;
	page: number;
	y: number;
}
interface Probe {
	groups: Record<string, Record<string, Row[]>>;
}
const probe = JSON.parse(
	readFileSync(
		new URL('./__fixtures__/word-table/table-probe.json', import.meta.url),
		'utf8',
	).replace(/^﻿/, ''),
) as Probe;
const measurer = wordProbeMeasurer;
const PT_PER_PX = 0.75;
// Word draws a 0.5 pt border one whole 0.75 pt pixel wide, so each bordered row edge can differ by up
// to 0.25 pt from the layout's 0.5 pt, and a three-row merge adds several (up to 1.6 pt measured);
// tables without borders stay within 1 pt.
const toleranceFor = (name: string) => (name.startsWith('vmn') ? 1 : 2);
// Documents whose difference is understood and not fixed; the test fails when one starts to match.
const knownGaps: Record<string, string> = {
	bd_dbl4: 'a double border is not a solid edge, so the layout reserves no space for it',
};

type Found = Map<string, { page: number; y: number }>;
function collect(
	blocks: LayoutBlockBox[],
	page: LayoutPageBox,
	originPx: number,
	found: Found,
): void {
	for (const box of blocks) {
		if (box.kind === 'paragraph') {
			const text = box.lines.flatMap((line) => line.fragments.map((f) => f.text)).join('');
			if (!text || found.has(text)) continue;
			found.set(text, {
				page: page.index + 1,
				y: (page.marginTopPx + originPx + box.yPx + (box.lines[0]?.yPx ?? 0)) * PT_PER_PX,
			});
		} else
			for (const row of box.rows)
				for (const cell of row.cells) collect(cell, page, originPx + box.yPx + row.yPx, found);
	}
}

const groups = Object.entries(probe.groups);

describe('table row sizing against the Word probe', () => {
	it('measured every document the generator builds', () => {
		const measured = groups.flatMap(([, docs]) => Object.keys(docs)).sort();
		expect(measured).toEqual(Object.keys(tableBodies()).sort());
	});
	describe.each(groups)('%s', (_group, docs) => {
		it.each(Object.keys(docs))('%s', async (name) => {
			const loaded = await loadDocx(await tableDocx(name));
			const found: Found = new Map();
			for (const page of layoutDocumentModel(loaded.model, measurer).pages)
				for (const column of page.columns) collect(column.blocks, page, 0, found);
			const wrong: string[] = [];
			for (const row of docs[name]!) {
				// Word reports an empty merged-away cell at the top of the merge, not in its own row.
				if (/e\d$/.test(row.label)) continue;
				const mine = found.get(row.label);
				if (!mine || mine.page !== row.page || Math.abs(mine.y - row.y) > toleranceFor(name))
					wrong.push(
						`${row.label}: ${mine ? `page ${mine.page} y ${mine.y.toFixed(1)}` : 'missing'}, Word ${row.page}/${row.y}`,
					);
			}
			if (name in knownGaps) expect(wrong, `${name} now matches: remove its gap`).not.toEqual([]);
			else expect(wrong).toEqual([]);
		});
	});
});
