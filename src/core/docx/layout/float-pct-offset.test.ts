import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer, createPageStories, layoutDocumentModel } from './index';
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

it('places a footer box from wp14:pctPosVOffset and from mc:AlternateContent (#408)', async () => {
	const top = async (name: string) => {
		const { stories } = await open(name);
		return at(at(stories(0), 0).floats, 0).yPx;
	};
	const control = await top('issue-408-absolute-offset-control');
	// 9854565 EMU: the top of the box is 776.0 pt.
	expect(control).toBeCloseTo(776.0 * PT, 0);
	expect(await top('issue-408-pct-offset-alternate-content')).toBeCloseTo(control, 0);
	expect(await top('issue-408-pct-offset-only')).toBeCloseTo(control, 0);
});
