import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer, createPageStories, layoutDocumentModel } from './index';
import type { LayoutBlockBox } from './result';
import { at } from './test-helpers';

const measurer = createFakeMeasurer();
const PT = 96 / 72;

async function open(name: string) {
	const { model } = await loadDocx(
		await readFile(new URL(`./fixtures/word-layout-reports/${name}.docx`, import.meta.url)),
	);
	const { pages } = layoutDocumentModel(model, measurer);
	return { model, pages, stories: createPageStories(model, pages, measurer) };
}

const lineTexts = (blocks: LayoutBlockBox[]) =>
	blocks.flatMap((block) =>
		block.kind === 'paragraph'
			? block.lines.map((line) => line.fragments.map((fragment) => fragment.text).join(''))
			: [],
	);

it('evaluates a PAGE field inside a footer text box on every page (#412)', async () => {
	const { pages, stories } = await open('issue-412-page-field-in-footer-text-box');
	expect(pages).toHaveLength(3);
	const shown = pages.map((_, index) => {
		const box = at(at(stories(index), 0).floats, 0);
		return lineTexts(box.textBoxLayout?.blocks ?? []);
	});
	expect(shown).toEqual([['1'], ['2'], ['3']]);
	const control = await open('issue-412-page-field-in-footer-paragraph');
	expect(control.pages.map((_, index) => lineTexts(at(control.stories(index), 0).blocks))).toEqual([
		['Footer 1'],
		['Footer 2'],
		['Footer 3'],
	]);
});
