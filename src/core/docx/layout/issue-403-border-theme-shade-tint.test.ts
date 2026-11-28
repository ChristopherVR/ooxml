import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer, layoutDocumentModel } from './index';

async function layout(name: string) {
	const { model } = await loadDocx(
		await readFile(new URL(`./fixtures/word-layout-reports/${name}.docx`, import.meta.url)),
	);
	return layoutDocumentModel(model, createFakeMeasurer());
}

const channels = (hex: string | undefined) =>
	[1, 3, 5].map((start) => Number.parseInt(hex?.slice(start, start + 2) ?? 'ff', 16));
const close = (actual: string | undefined, expected: string) =>
	channels(actual).every((value, index) => Math.abs(value - (channels(expected)[index] ?? 0)) <= 1);

/** The top border colour and the paragraph's text top (pt) for each bordered paragraph. */
async function borderColors(name: string) {
	const result = await layout(name);
	const found: { color: string | undefined; y: number }[] = [];
	for (const page of result.pages)
		for (const column of page.columns)
			for (const block of column.blocks)
				if (block.kind === 'paragraph' && block.frame?.borders?.top)
					found.push({
						color: block.frame.borders.top.color,
						y: (page.marginTopPx + block.yPx + (block.lines[0]?.yPx ?? 0)) * 0.75,
					});
	return found;
}

it('applies themeShade and themeTint to paragraph borders and keeps both (#403)', async () => {
	const found = await borderColors('issue-403-border-theme-shade-tint');
	expect(found).toHaveLength(2);
	// The shared HSL tint lands within one unit per channel of Word's #E5B8B7.
	expect(close(found[0]?.color, '#622423')).toBe(true);
	expect(close(found[1]?.color, '#E5B8B7')).toBe(true);
});

it('draws plain border colours as they are (#403 control)', async () => {
	const plain = await borderColors('issue-403-border-plain-colors');
	const themed = await borderColors('issue-403-border-theme-shade-tint');
	expect(plain.map((item) => item.color?.toUpperCase())).toEqual(['#622423', '#E5B8B7']);
	expect(themed.map((item) => item.y)).toEqual(plain.map((item) => item.y));
});
