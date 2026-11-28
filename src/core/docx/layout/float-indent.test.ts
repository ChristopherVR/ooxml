import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { signedTwips, twips } from '../../units';
import type { Paragraph } from '../model';
import { createFakeMeasurer, layoutDocumentModel } from './index';
import { at } from './test-helpers';

// Issue #370: a paragraph indent is measured from the margin, not added to the edge of a
// square-wrapped float beside it. The picture is 3165 twips wide at the column's left edge
// (plus 12px of side distance), so text at an indent below that edge starts at the float's edge.
it('starts indented text at the float edge when the indent is smaller than it', async () => {
	const { model } = await loadDocx(
		await readFile(new URL('./fixtures/word-parity/issue-370-float-indent.docx', import.meta.url)),
	);
	const { pages } = layoutDocumentModel(model, createFakeMeasurer());
	const page = at(pages, 0);
	const picture = at(page.floats ?? [], 0);
	const edge = picture.xPx + picture.widthPx + 12 - page.marginLeftPx;
	const firstXs = at(page.columns, 0)
		.blocks.filter((block) => block.kind === 'paragraph')
		.map((block) => at(block.lines, 0).fragments[0]?.xPx ?? 0);
	// "Anchor paragraph.", indent 4820 (beyond the float), indent 1440 (inside it), hanging 357.
	const indent4820 = (4820 / 20) * (96 / 72);
	expect(firstXs[1]).toBeCloseTo(indent4820);
	expect(firstXs[2]).toBeCloseTo(edge);
	expect(firstXs[3]).toBeCloseTo(edge);
});

// A negative indent reaches into the margin whenever no float is beside the paragraph.
it('keeps negative indents in the margin without a float', async () => {
	const { model } = await loadDocx(
		await readFile(new URL('./fixtures/word-parity/issue-370-float-indent.docx', import.meta.url)),
	);
	const source = model.blocks[3] as Paragraph;
	const make = (id: string, indent: Partial<Paragraph>) =>
		({
			...structuredClone(source),
			id,
			indentLeftTwips: undefined,
			hangingTwips: undefined,
			firstLineTwips: undefined,
			...indent,
		}) as Paragraph;
	// No float: the picture's paragraph is left out.
	model.blocks = [
		make('neg-left', { indentLeftTwips: signedTwips(-720) }),
		make('neg-hang', { indentLeftTwips: signedTwips(0), hangingTwips: twips(360) }),
		make('neg-right', { indentRightTwips: signedTwips(-720) }),
	] as never;
	const page = at(layoutDocumentModel(model, createFakeMeasurer()).pages, 0);
	const [a, b, c] = at(page.columns, 0).blocks.map((block) =>
		block.kind === 'paragraph' ? at(block.lines, 0) : undefined,
	);
	expect(a?.fragments[0]?.xPx).toBeCloseTo(-48);
	expect(b?.fragments[0]?.xPx).toBeCloseTo(-24);
	expect(c?.fragments[0]?.xPx).toBeCloseTo(0);
	// A negative right indent lets the line run past the margin (604.8px wide column).
	const end = Math.max(...(c?.fragments ?? []).map((fragment) => fragment.xPx + fragment.widthPx));
	expect(end).toBeGreaterThan(604.8);
});
