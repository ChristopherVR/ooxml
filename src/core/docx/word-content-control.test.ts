// Line heights of lines holding content controls against what desktop Word measured
// (`__fixtures__/word-content-control/line-probe.json`, written by
// `scripts/docx-word-content-control-probe.ps1`).
//
// What the probe showed (the corpus gap `content-control-line-height`): a content control does not
// change its line. A check box content control and a plain MS Gothic run give the same 14.25 pt line
// at 11 pt (against 12 to 12.75 pt for Courier New), so the line height follows the FONT of the
// symbol, which the host's measurer owns. `word-measurer.ts` carries the measured heights; the
// layout only has to ask for the font of each run.
//
// `unmodelled` lists the case whose height is a font fallback that depends on the installed fonts:
// a symbol in a Courier New run inside a check box (15.75 pt, against 12 pt for the same symbol in a
// plain run), which no per-font height can reproduce. A Latin run set in MS Gothic is 15 pt in Word
// and 14.25 pt here, inside the 1 pt tolerance.
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { loadDocx } from './index';
import { layoutDocumentModel } from './layout/layout';
import { wordProbeMeasurer } from './word-measurer';
import { contentControlCases, contentControlDocx } from './word-content-control-docs';

interface Probe {
	cases: string[];
	rows: { label: string; y: number }[];
}
const probe = JSON.parse(
	readFileSync(
		new URL('./__fixtures__/word-content-control/line-probe.json', import.meta.url),
		'utf8',
	).replace(/^﻿/, ''),
) as Probe;
const PT_PER_PX = 0.75;
const unmodelled = new Set(['cbcourier']);

/** Top y (pt) of every paragraph, in order. */
async function tops(): Promise<number[]> {
	const loaded = await loadDocx(await contentControlDocx());
	const out: number[] = [];
	for (const page of layoutDocumentModel(loaded.model, wordProbeMeasurer).pages)
		for (const column of page.columns)
			for (const box of column.blocks)
				if (box.kind === 'paragraph')
					out.push((page.marginTopPx + box.yPx + (box.lines[0]?.yPx ?? 0)) * PT_PER_PX);
	return out;
}

describe('content control line heights against the Word probe', () => {
	it('measured every case the generator builds', () => {
		expect(probe.cases).toEqual(Object.keys(contentControlCases));
	});
	it.each(Object.keys(contentControlCases))('%s', async (name) => {
		const at = probe.cases.indexOf(name);
		const ours = await tops();
		const wordHeight = probe.rows[2 * at + 2]!.y - probe.rows[2 * at + 1]!.y;
		const ourHeight = ours[2 * at + 2]! - ours[2 * at + 1]!;
		if (unmodelled.has(name)) expect(Math.abs(ourHeight - wordHeight)).toBeGreaterThan(1);
		else expect(Math.abs(ourHeight - wordHeight)).toBeLessThanOrEqual(1);
	});
});
