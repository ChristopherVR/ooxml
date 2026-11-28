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

it('keeps the space after the last header and footer paragraph inside them (#401)', async () => {
	const bug = await open('issue-401-header-footer-space-after');
	const control = await open('issue-401-header-footer-no-space-after');
	const bugPage = at(bug.pages, 0);
	const controlPage = at(control.pages, 0);
	// Word: the body starts 10 pt lower and the footer lines 10 pt higher.
	expect(bugPage.marginTopPx - controlPage.marginTopPx).toBeCloseTo(10 * PT, 1);
	const header = at(bug.stories(0), 0);
	const footer = at(bug.stories(0), 1);
	const controlFooter = at(control.stories(0), 1);
	expect(header.heightPx - at(control.stories(0), 0).heightPx).toBeCloseTo(10 * PT, 1);
	expect(controlFooter.y - footer.y).toBeCloseTo(10 * PT, 1);
	expect(bugPage.marginBottomPx).toBeCloseTo(controlPage.marginBottomPx, 1);
});
