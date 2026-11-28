// @vitest-environment node
import { readFile } from 'node:fs/promises';
import { loadDocx, type Paragraph } from 'ooxml-core/docx';
import { expect, it } from 'vitest';
import { resolvedNodeTabStops, resolvedTabStops } from './resolved-tab-stops';

// Issue #414: the stops of a paragraph style show and apply, merged with the direct ones.
const fixture = new URL(
	'../../../core/docx/layout/fixtures/word-layout-reports/issue-414-style-tab-stops.docx',
	import.meta.url,
);
const positions = (stops: { posTwips: number }[]) => stops.map((stop) => Number(stop.posTwips));
const stop = (posTwips: number, align: 'left' | 'right' | 'clear') =>
	({ posTwips, align }) as never;

it('takes the stops of the style and of the style it is based on', async () => {
	const { model } = await loadDocx(new Uint8Array(await readFile(fixture)));
	const paragraphs = model.blocks.filter((block): block is Paragraph => block.type === 'paragraph');
	expect(positions(resolvedTabStops(paragraphs[0]!, model))).toEqual([4536, 9070]);
	// "Derived" clears 4536 and adds 2000 on top of its parent's stops.
	expect(positions(resolvedTabStops(paragraphs[1]!, model))).toEqual([2000, 9070]);
});

it('merges direct stops over the style, a clear stop removing an inherited one', async () => {
	const { model } = await loadDocx(new Uint8Array(await readFile(fixture)));
	const direct = [stop(9070, 'clear'), stop(3000, 'left')];
	expect(positions(resolvedTabStops({ style: 'TabbedBody', tabStops: direct }, model))).toEqual([
		3000, 4536,
	]);
	// A node's attributes resolve the same way; its own stops stay untouched.
	const attrs = { style: 'TabbedBody', tabStops: direct };
	expect(positions(resolvedNodeTabStops(attrs, model))).toEqual([3000, 4536]);
	expect(attrs.tabStops).toEqual(direct);
	// Without a catalog only the direct stops count, minus the clear ones.
	expect(positions(resolvedNodeTabStops(attrs, undefined))).toEqual([3000]);
});
