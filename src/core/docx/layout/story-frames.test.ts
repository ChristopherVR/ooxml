import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { loadDocx, twips } from '../index';
import { createFakeMeasurer, createPageStories, layoutDocumentModel } from './index';
import { at } from './test-helpers';

const measure = createFakeMeasurer();

async function load(name: string) {
	const url = new URL(`./fixtures/story-frames/issue-350-${name}.docx`, import.meta.url);
	return (await loadDocx(await readFile(url))).model;
}

it('positions the footer frame from the page origin on repeated calls and pages', async () => {
	const model = await load('footer-frame-page-anchor');
	const { pages } = layoutDocumentModel(model, measure);
	const page = at(pages, 0);
	const stories = createPageStories(model, [page, { ...page, index: 1 }], measure);
	for (const index of [0, 1, 0]) {
		const footer = at(stories(index), 0);
		expect(footer.kind).toBe('footer');
		const frame = at(
			footer.blocks.filter((box) => box.kind === 'paragraph' && box.floating),
			0,
		);
		expect(footer.x + (frame.xPx ?? 0)).toBeCloseTo(10000 / 15);
		expect(footer.y + frame.yPx).toBeCloseTo(15900 / 15);
		expect(footer.heightPx).toBeCloseTo(18.2, 1);
	}
});

it('keeps the unframed footer in flow', async () => {
	const model = await load('footer-no-frame-control');
	const { pages } = layoutDocumentModel(model, measure);
	const footer = at(createPageStories(model, pages, measure)(0), 0);
	expect(footer.blocks).toHaveLength(2);
	expect(footer.blocks.every((box) => box.kind === 'paragraph' && !box.floating)).toBe(true);
	expect(footer.heightPx).toBeCloseTo(47.07, 1);
});

it('resolves page and margin alignment against each physical page', async () => {
	const model = await load('footer-frame-page-anchor');
	const section = at(model.sections, 0);
	const content = section.footers?.default;
	if (!content) throw new Error('Expected a default footer');
	const paragraph = at(
		content.blocks.filter((block) => block.type === 'paragraph').filter((block) => block.frame),
		0,
	);
	paragraph.frame = {
		...paragraph.frame,
		hAnchor: 'page',
		xAlign: 'right',
		vAnchor: 'margin',
		yAlign: 'bottom',
		yTwips: twips(0),
	};
	const page = at(layoutDocumentModel(model, measure).pages, 0);
	const taller = {
		...page,
		index: 1,
		heightPx: page.heightPx + 100,
		marginBottomPx: page.marginBottomPx + 30,
	};
	const stories = createPageStories(model, [page, taller], measure);
	for (const [index, physical] of [page, taller].entries()) {
		const footer = at(stories(index), 0);
		const frame = at(
			footer.blocks.filter((box) => box.kind === 'paragraph' && box.floating),
			0,
		);
		expect(footer.x + (frame.xPx ?? 0) + (frame.widthPx ?? 0)).toBeCloseTo(physical.widthPx);
		expect(
			footer.y +
				frame.yPx +
				(frame.kind === 'paragraph' ? (frame.floating?.frameHeightPx ?? 0) : 0),
		).toBeCloseTo(physical.heightPx - physical.marginBottomPx);
	}
});

it('keeps the body control frame at its page offset', async () => {
	const model = await load('body-frame-page-anchor');
	const page = at(layoutDocumentModel(model, measure).pages, 0);
	const column = at(page.columns, 0);
	const frame = at(
		column.blocks.filter((box) => box.kind === 'paragraph' && box.floating),
		0,
	);
	expect(page.marginLeftPx + column.xPx + (frame.xPx ?? 0)).toBeCloseTo(10000 / 15);
	expect(page.marginTopPx + frame.yPx).toBeCloseTo(15900 / 15);
});
