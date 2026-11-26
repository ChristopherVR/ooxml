import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer, createPageStories, layoutDocumentModel } from './index';
import type { LayoutPageBox } from './result';
import { at } from './test-helpers';
import { floatExclusions } from './wrap';

const measure = createFakeMeasurer();
const load = async (name: string) =>
	(
		await loadDocx(
			await readFile(
				new URL(`./fixtures/story-float-wrap/issue-351-${name}.docx`, import.meta.url),
			),
		)
	).model;

function bodyLines(page: LayoutPageBox) {
	return at(page.columns, 0).blocks.flatMap((box) =>
		box.kind === 'paragraph'
			? box.lines
					.filter((line) => line.fragments.some((fragment) => fragment.text.trim()))
					.map((line) => ({
						...line,
						yPx: page.marginTopPx + box.yPx + line.yPx,
					}))
			: [],
	);
}

it('clears the header text box on every body line beside it', async () => {
	const model = await load('header-shape-wrap-square');
	const { pages } = layoutDocumentModel(model, measure);
	const stories = createPageStories(model, pages, measure);
	for (const [index, page] of pages.entries()) {
		const picture = at(at(stories(index), 0).floats, 0);
		expect(picture).toMatchObject({ wrap: 'square', widthPx: 211, heightPx: 854 });
		const bottom = picture.yPx + picture.heightPx;
		const beside = bodyLines(page).filter(
			(line) => line.yPx < bottom && line.yPx + line.heightPx > picture.yPx,
		);
		expect(beside.length).toBeGreaterThan(0);
		const inset = picture.xPx + picture.widthPx + 12 - page.marginLeftPx;
		expect(inset).toBeCloseTo(142);
		for (const line of beside) {
			expect(at(line.fragments, 0).xPx).toBeCloseTo(inset);
			expect(line.fragments.every((fragment) => fragment.xPx >= inset - 0.01)).toBe(true);
		}
	}
	const first = at(pages, 0);
	const picture = at(at(stories(0), 0).floats, 0);
	const below = bodyLines(first).filter((line) => line.yPx >= picture.yPx + picture.heightPx);
	expect(below.length).toBeGreaterThan(0);
	expect(below.map((line) => at(line.fragments, 0).xPx)).toEqual(below.map(() => 0));
	expect(layoutDocumentModel(model, measure).pages.map(bodyLines)).toEqual(pages.map(bodyLines));
});

it('keeps the body-anchored control wrapping at the same inset', async () => {
	const model = await load('header-shape-wrap-square-control');
	const { pages } = layoutDocumentModel(model, measure);
	expect(pages).toHaveLength(2);
	const first = at(pages, 0);
	expect(at(first.floats, 0).wrap).toBe('square');
	expect(bodyLines(first).some((line) => Math.abs(at(line.fragments, 0).xPx - 142) < 0.01)).toBe(
		true,
	);
});

it('lets text use the full width beside a nonwrapping header float', async () => {
	const model = await load('header-shape-wrap-square');
	const header = at(model.sections, 0).headers?.default;
	const run = header?.blocks
		.flatMap((block) => (block.type === 'paragraph' ? block.runs : []))
		.find((run) => run.image?.anchored);
	if (!run?.image?.placement) throw new Error('Expected a header float');
	run.image.placement = { ...run.image.placement, wrap: 'none', behindText: true };
	const { pages } = layoutDocumentModel(model, measure);
	const page = at(pages, 0);
	const picture = at(at(createPageStories(model, pages, measure)(0), 0).floats, 0);
	expect(picture).toMatchObject({ wrap: 'none', behindText: true });
	const overlapping = bodyLines(page).filter(
		(line) => line.yPx >= picture.yPx && line.yPx < picture.yPx + picture.heightPx,
	);
	expect(overlapping.length).toBeGreaterThan(0);
	expect(overlapping.every((line) => at(line.fragments, 0).xPx === 0)).toBe(true);
});

it('wraps body lines around the same shape anchored in a footer', async () => {
	const model = await load('header-shape-wrap-square');
	const section = at(model.sections, 0);
	const content = section.headers?.default;
	const run = content?.blocks
		.flatMap((block) => (block.type === 'paragraph' ? block.runs : []))
		.find((run) => run.image?.anchored);
	if (!content || !run?.image?.placement) throw new Error('Expected a header float');
	delete section.headers;
	section.footers = { default: content };
	run.image.placement = { ...run.image.placement, relativeFromV: 'page', offsetYPx: 300 };
	const { pages } = layoutDocumentModel(model, measure);
	const page = at(pages, 0);
	const footer = at(createPageStories(model, pages, measure)(0), 0);
	expect(footer.kind).toBe('footer');
	const picture = at(footer.floats, 0);
	expect(picture.yPx).toBe(300);
	const lines = bodyLines(page);
	expect(at(at(lines, 0).fragments, 0).xPx).toBe(0);
	const beside = lines.filter((line) => line.yPx >= picture.yPx);
	expect(beside.length).toBeGreaterThan(0);
	expect(beside.every((line) => Math.abs(at(line.fragments, 0).xPx - 142) < 0.01)).toBe(true);
});

it('keeps wrapping on pages that only exist after reflow', async () => {
	const model = await load('header-shape-wrap-square');
	const run = at(model.sections, 0)
		.headers?.default?.blocks.flatMap((block) => (block.type === 'paragraph' ? block.runs : []))
		.find((run) => run.image?.anchored);
	if (!run?.image) throw new Error('Expected a header float');
	run.image.widthPx = 580;
	run.image.heightPx = 950;
	const { pages } = layoutDocumentModel(model, measure);
	expect(pages.length).toBeGreaterThan(4);
	const stories = createPageStories(model, pages, measure);
	for (const [index, page] of pages.entries()) {
		const { xPx, widthPx } = at(at(stories(index), 0).floats, 0);
		const inset = xPx + widthPx + 12 - page.marginLeftPx;
		const lines = bodyLines(page);
		expect(lines.length).toBeGreaterThan(0);
		expect(lines.every((line) => at(line.fragments, 0).xPx >= inset - 0.01)).toBe(true);
	}
});

it('turns only wrapping page floats into exclusions with side clearance', async () => {
	const model = await load('header-shape-wrap-square');
	const { pages } = layoutDocumentModel(model, measure);
	const picture = at(at(createPageStories(model, pages, measure)(0), 0).floats, 0);
	expect(
		floatExclusions([picture, { ...picture, inCell: true }, { ...picture, wrap: 'none' }]),
	).toEqual([
		{
			xPx: picture.xPx - 12,
			yPx: picture.yPx,
			widthPx: picture.widthPx + 24,
			heightPx: picture.heightPx,
			topAndBottom: false,
			bothSides: true,
		},
	]);
});
