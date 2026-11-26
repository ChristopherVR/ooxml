import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer, createPageStories, layoutDocumentModel } from './index';
import { at } from './test-helpers';

async function load(name: string) {
	const url = new URL(`./fixtures/vml-floats/issue-345-${name}.docx`, import.meta.url);
	const bytes = await readFile(url);
	return { loaded: await loadDocx(bytes), bytes };
}

it.each([
	['vml-absolute-picture-in-header', 'center'],
	['vml-absolute-picture-offset', 'offset'],
])('floats %s behind the body without enlarging the header', async (name, position) => {
	const {
		loaded: { model },
	} = await load(name);
	const measure = createFakeMeasurer();
	const { pages } = layoutDocumentModel(model, measure);
	expect(pages).toHaveLength(3);
	const stories = createPageStories(model, pages, measure);
	for (const [index, page] of pages.entries()) {
		expect(page.marginTopPx).toBeCloseTo(1417 / 15);
		const header = at(stories(index), 0);
		expect(header.kind).toBe('header');
		expect(header.floats).toHaveLength(1);
		const picture = at(header.floats, 0);
		expect(picture).toMatchObject({ behindText: true, wrap: 'none' });
		expect(picture.widthPx).toBeCloseTo((453.35 * 4) / 3);
		expect(picture.heightPx).toBeCloseTo((442.45 * 4) / 3);
		expect(picture.xPx).toBeCloseTo(
			page.marginLeftPx +
				(position === 'center' ? (header.widthPx - picture.widthPx) / 2 : (-0.6 * 4) / 3),
		);
		expect(picture.yPx).toBeCloseTo(
			page.marginTopPx +
				(position === 'center'
					? (page.heightPx - page.marginTopPx - page.marginBottomPx - picture.heightPx) / 2
					: (133.4 * 4) / 3),
		);
	}
});

it('keeps the picture watermark control on its dedicated path', async () => {
	const {
		loaded: { model },
	} = await load('vml-watermark-id-control');
	const measure = createFakeMeasurer();
	const { pages } = layoutDocumentModel(model, measure);
	expect(pages).toHaveLength(3);
	const header = at(createPageStories(model, pages, measure)(0), 0);
	expect(header.floats).toEqual([]);
	const runs = header.source.flatMap((block) => (block.type === 'paragraph' ? block.runs : []));
	expect(runs.find((run) => run.image)?.image).toMatchObject({
		widthPx: 0,
		heightPx: 0,
		pictureWatermark: { kind: 'picture', washout: true },
	});
});

it('preserves source VML when the header paragraph is edited', async () => {
	const { loaded, bytes } = await load('vml-absolute-picture-offset');
	const model = structuredClone(loaded.model);
	const header = model.sections?.[0]?.headers?.default;
	const paragraph = at(
		header?.blocks
			.filter((block) => block.type === 'paragraph')
			.filter((block) => block.runs.some((run) => run.image)),
		0,
	);
	paragraph.runs.push({ text: 'New header text' });
	const headerXml = async (doc: Uint8Array) =>
		(await JSZip.loadAsync(doc)).file('word/header1.xml')?.async('string');
	const source = await headerXml(bytes);
	const saved = await headerXml(await loaded.save(model));
	expect(saved).toContain('New header text');
	const shape = (xml: string | undefined) => xml?.match(/<v:shape\b[^]*?<\/v:shape>/)?.[0];
	expect(shape(source)).toBeDefined();
	expect(shape(saved)).toBe(shape(source));
	const run = at(
		paragraph.runs.filter((run) => run.image),
		0,
	);
	if (!run.image?.placement) throw new Error('Expected a VML float');
	run.image.placement = { ...run.image.placement, offsetXPx: 200, wrap: 'square' };
	expect(shape(await headerXml(await loaded.save(model)))).toBe(shape(source));
});
