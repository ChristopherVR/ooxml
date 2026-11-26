import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer, layoutDocumentModel } from './index';
import { paragraphFrameBounds } from './result';
import { layoutRow, splitRowAtHeight } from './table-layout';
import { at } from './test-helpers';

async function layout(name: string) {
	const url = new URL(
		`./fixtures/float-in-bordered-paragraph/issue-352-${name}.docx`,
		import.meta.url,
	);
	const { model } = await loadDocx(await readFile(url));
	const { pages } = layoutDocumentModel(model, createFakeMeasurer());
	const page = at(pages, 0);
	const box = at(
		at(page.columns, 0).blocks.filter((block) => block.kind === 'paragraph'),
		0,
	);
	return { page, box };
}

it.each([
	['picture-in-bordered-paragraph', true],
	['picture-in-bordered-paragraph-in-front', false],
])('frames only the text below %s', async (name, behindText) => {
	const { page, box } = await layout(name);
	const picture = at(page.floats, 0);
	expect(picture).toMatchObject({ wrap: 'topAndBottom', behindText });
	expect(box.lines).toHaveLength(2);
	const gap = at(box.lines, 0).gapBeforePx ?? 0;
	expect(gap).toBeGreaterThan(150);
	expect(box.frame).toMatchObject({
		topPx: gap,
		borders: { top: expect.anything() },
		shading: expect.any(String),
	});
	if (!box.frame) throw new Error('Expected a paragraph frame');
	const bounds = paragraphFrameBounds(box.frame, box.heightPx);
	expect(bounds.heightPx).toBeCloseTo(56.95, 1);
	expect(bounds.topPx + bounds.heightPx).toBeCloseTo(box.heightPx);
	expect(picture.yPx + picture.heightPx).toBeLessThanOrEqual(
		page.marginTopPx + box.yPx + bounds.topPx + 0.01,
	);
});

it('retains the plain paragraph’s leading wrap band and text geometry', async () => {
	const { box } = await layout('picture-in-plain-paragraph-control');
	expect(box.frame).toBeUndefined();
	expect(box.heightPx).toBeCloseTo(238.95, 1);
	expect(box.lines).toHaveLength(2);
	expect(at(box.lines, 0).yPx).toBeCloseTo(186, 1);
	expect(at(box.lines, 1).yPx).toBeCloseTo(212.5, 1);
});

it('excludes a cell float from shading and clears the offset on the continuation', () => {
	const row = layoutRow(
		{
			cells: [
				{
					paragraphs: [
						{
							kind: 'paragraph',
							id: 'shaded',
							shading: '#e0e0e0',
							runs: [{ text: 'Words in the shaded cell. '.repeat(8) }],
							floats: [
								{
									partName: 'word/media/image.png',
									contentType: 'image/png',
									widthPx: 50,
									heightPx: 40,
									wrap: 'topAndBottom',
								},
							],
						},
					],
				},
			],
		},
		200,
		createFakeMeasurer(),
		() => {},
	);
	const paragraph = at(at(row.cells, 0), 0);
	const firstLine = at(paragraph.lines, 0);
	expect(paragraph.frame).toMatchObject({ topPx: 40, shading: '#e0e0e0' });
	const { before, after } = splitRowAtHeight(
		row,
		paragraph.yPx + firstLine.yPx + firstLine.heightPx,
	);
	expect(at(at(before.cells, 0), 0).frame?.topPx).toBe(40);
	expect(at(at(after?.cells, 0), 0).frame?.topPx).toBe(0);
});
