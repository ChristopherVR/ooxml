import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer, layoutDocumentModel } from './index';
import { at } from './test-helpers';

const measurer = createFakeMeasurer();
const PT = 96 / 72;

async function open(name: string) {
	const { model } = await loadDocx(
		await readFile(new URL(`./fixtures/word-layout-reports/${name}.docx`, import.meta.url)),
	);
	const { pages } = layoutDocumentModel(model, measurer);
	return { pages };
}

/** Where text beside the square-wrapped box starts on its first line, in page pixels. */
async function textStart(name: string) {
	const { pages } = await open(name);
	const page = at(pages, 0);
	const box = at(page.floats, 0);
	const first = at(at(page.columns, 0).blocks, 0);
	if (first.kind !== 'paragraph') throw new Error('Expected a paragraph');
	const line = at(first.lines, 0);
	return { x: page.marginLeftPx + at(line.fragments, 0).xPx, boxRight: box.xPx + box.widthPx };
}

it('wraps text at the object edge plus effectExtent plus distR (#416)', async () => {
	const control = await textStart('issue-416-default-wrap-control');
	// Word: 9 pt default distance from the box edge (238.1 pt).
	expect(control.x - control.boxRight).toBeCloseTo(9 * PT, 1);
	const effect = await textStart('issue-416-effect-extent-wrap');
	// 0.75 pt effectExtent: Word starts the text at 238.85 pt.
	expect(effect.x - effect.boxRight).toBeCloseTo((9 + 0.75) * PT, 1);
	const distance = await textStart('issue-416-wrap-distance-36pt');
	// distR 36 pt plus a 3 pt effectExtent: Word starts the text at 268.1 pt.
	expect(distance.x - distance.boxRight).toBeCloseTo(39 * PT, 1);
});
