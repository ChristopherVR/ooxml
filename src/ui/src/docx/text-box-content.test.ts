// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { loadDocx, type DocumentModel, type Paragraph } from 'ooxml-core/docx';
import {
	createFakeMeasurer,
	layoutDocumentModel,
	type LayoutPageBox,
} from 'ooxml-core/docx/layout';
import { letterheadFixture } from '../../../core/docx/test-support/letterhead-fixture';
import { buildHeaderElement } from './header-footer-view';
import { renderBlocks } from './story-render';
import { docToModel, modelToDoc } from './model-adapter';
import { decoratePages } from './print-header-footer';
import { renderPrintLayout } from './print-layout';
import { shapeNodeView } from './shape-node-view';
import { must } from './test-support';

const firstHeader = (model: DocumentModel) =>
	must(model.sections?.[0]?.headers?.first, 'first-page header');
const headerXml = async (bytes: Uint8Array) =>
	(await (await JSZip.loadAsync(bytes)).file('word/header1.xml')!.async('string')) as string;

/** The letterhead page: A4 with Word's margins from the fixture. */
const firstPage: LayoutPageBox = {
	index: 0,
	sectionIndex: 0,
	pageInSection: 0,
	widthPx: 794,
	heightPx: 1123,
	marginTopPx: 96,
	marginRightPx: 96,
	marginBottomPx: 96,
	marginLeftPx: 227,
	columns: [],
};

/** An editing surface over `blocks` with the editor's shape node view. */
function editingSurface(model: DocumentModel, blocks: DocumentModel['blocks']): EditorView {
	return new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc: modelToDoc({ ...model, blocks }) }),
		nodeViews: {
			image: (node, nodeView, getPos) =>
				node.attrs.shape
					? shapeNodeView(
							node,
							nodeView,
							getPos as () => number,
							() => model.theme,
							() => model,
						)
					: (undefined as never),
		},
	});
}

/** What a text box and a line in a story must show: the box's paragraphs and two drawn shapes. */
function expectDrawnLetterhead(root: HTMLElement) {
	expectDrawnShapes(root);
	const content = must(root.querySelector<HTMLElement>('.dve-text-box-content'), 'text box');
	const paragraphs = [...content.querySelectorAll<HTMLElement>(':scope > p')];
	expect(paragraphs.map((paragraph) => paragraph.textContent)).toEqual([
		'First Partner',
		'Specialist in contract law',
		'',
		'12 Example Street',
	]);
	const name = must(paragraphs[0]?.querySelector('strong, [style*="font-weight"]'), 'bold run');
	expect(name.textContent).toBe('First Partner');
	expect(paragraphs[1]?.innerHTML).toMatch(/font-size:\s*9pt/);
	expect(paragraphs[2]?.style.borderBottom).toMatch(/solid/);
	expect(content.style.padding).toBe('4.8px 9.6px');
}

/** Both shapes drawn, with nothing left as a placeholder. */
function expectDrawnShapes(root: HTMLElement) {
	expect(root.querySelector('.dve-print-picture-missing')).toBeNull();
	expect(root.querySelector('.dve-image-placeholder')).toBeNull();
	expect(root.textContent).not.toContain('Shape');
	const strokes = [...root.querySelectorAll('svg path')].map((path) => path.getAttribute('stroke'));
	expect(strokes).toContain('#2E75B6');
}

/** The letterhead text box on a Print Layout page: the lines the page engine laid out. */
function expectPrintedLetterhead(root: HTMLElement) {
	expectDrawnShapes(root);
	expect(root.querySelector('.dve-text-box-content')).toBeNull();
	const layer = must(root.querySelector<HTMLElement>('.dve-print-text-box'), 'text box');
	const lines = [...layer.querySelectorAll<HTMLElement>('.dve-print-line')];
	expect(lines.map((line) => line.textContent)).toEqual([
		'First Partner',
		'Specialist in contract law',
		'',
		'12 Example Street',
	]);
	expect(lines[0]?.querySelector<HTMLElement>('span')?.style.fontWeight).toBe('700');
	expect(lines[1]?.querySelector<HTMLElement>('span')?.style.fontSize).toBe('9pt');
	expect(
		layer.querySelector<HTMLElement>('.dve-print-paragraph-frame')?.style.borderBottom,
	).toMatch(/solid/);
	// The insets; the box has no outline to keep clear of.
	expect(layer.style).toMatchObject({ left: '9.6px', top: '4.8px' });
}

describe('text boxes and lines in headers', () => {
	it('draws them with their paragraphs in story content', async () => {
		const { model } = await loadDocx(await letterheadFixture());
		const story = document.createElement('div');
		story.append(renderBlocks(firstHeader(model).blocks, model));
		expectDrawnLetterhead(story);
		expect(story.querySelectorAll('.dve-shape svg')).toHaveLength(2);
	});

	it('leaves them out of the Draft view header, as Word does, without the word "Shape"', async () => {
		const { model } = await loadDocx(await letterheadFixture());
		const header = must(buildHeaderElement(model, 'en'), 'header story');
		expect(header.querySelector('.dve-shape')).toBeNull();
		expect(header.querySelector('.dve-image-placeholder')).toBeNull();
		expect(header.textContent).not.toContain('Shape');
		expect(header.textContent).toContain('Letterhead');
		expect(header.querySelectorAll('.dve-floating-omitted')).toHaveLength(2);
	});

	it('draws them on the Print Layout page instead of empty placeholder boxes', async () => {
		const { model } = await loadDocx(await letterheadFixture());
		const sheet = document.createElement('div');
		decoratePages(model, [firstPage], [sheet]);
		expectPrintedLetterhead(sheet);
		const floats = [...sheet.querySelectorAll<HTMLElement>('.dve-print-float')];
		expect(floats).toHaveLength(2);
		// The text box sits where its page offsets put it, at its extent.
		expect(floats[0]?.style).toMatchObject({ left: '34px', top: '230px', width: '213px' });
	});

	it('draws a body text box with its paragraphs on Print Layout pages', async () => {
		const { model } = await loadDocx(await letterheadFixture());
		const story = firstHeader(model).blocks[0] as Paragraph;
		model.blocks = [{ ...story, id: 'body-box' }, ...model.blocks];
		const handle = renderPrintLayout(layoutDocumentModel(model, createFakeMeasurer()), undefined, {
			model,
		});
		const page = must(handle.element.querySelector<HTMLElement>('.dve-print-page'), 'page');
		const floats = page.querySelectorAll('.dve-print-float');
		expect(floats).toHaveLength(2);
		expect(page.querySelector('.dve-print-picture-missing')).toBeNull();
		expect(must(floats[0]?.querySelector('.dve-print-text-box'), 'box').textContent).toContain(
			'First Partner',
		);
	});

	it('draws the paragraphs over the shape on the editing surface', async () => {
		const { model } = await loadDocx(await letterheadFixture());
		const view = editingSurface(model, firstHeader(model).blocks);
		expectDrawnLetterhead(view.dom);
		view.destroy();
	});

	it('saves the header part unchanged after a round trip through the editor', async () => {
		const bytes = await letterheadFixture();
		const loaded = await loadDocx(bytes);
		const model = structuredClone(loaded.model);
		const header = firstHeader(model);
		const story = { ...model, blocks: header.blocks };
		header.blocks = docToModel(modelToDoc(story), story).blocks;
		expect(await headerXml(await loaded.save(model))).toBe(await headerXml(bytes));
	});
});

/** Word's own file: a red rectangle whose text, "Red box", sets no colour. */
const wordShapes = async () =>
	loadDocx(new Uint8Array(await readFile(resolve('../core/docx/__fixtures__/shapes/shapes.docx'))));

/** The colour the box's text is drawn in: its own, or the first ancestor's up to the box. */
function textColour(root: HTMLElement, text: string): string {
	const box = must(
		[...root.querySelectorAll<HTMLElement>('.dve-text-box-content, .dve-print-text-box')].find(
			(element) => element.textContent?.includes(text),
		),
		`text box with ${text}`,
	);
	const walker = document.createTreeWalker(box, NodeFilter.SHOW_TEXT);
	// Print Layout draws a line's words as separate fragments: find the first word's.
	const word = text.split(' ')[0]!;
	const texts: Node[] = [];
	for (let next = walker.nextNode(); next; next = walker.nextNode()) texts.push(next);
	const node = texts.find((item) => item.textContent?.includes(word));
	for (let element = node?.parentElement; element; element = element.parentElement) {
		if (element.style.color) return element.style.color;
		if (element === box) break;
	}
	return '';
}

describe('text box text colour', () => {
	const white = 'rgb(255, 255, 255)';
	const black = 'rgb(0, 0, 0)';
	const navy = 'rgb(31, 56, 100)';

	it('is white on a dark fill and keeps a run colour on the editing surface', async () => {
		const { model } = await wordShapes();
		const view = editingSurface(model, model.blocks);
		expect(textColour(view.dom, 'Red box')).toBe(white);
		view.destroy();
		const letterhead = (await loadDocx(await letterheadFixture())).model;
		const header = editingSurface(letterhead, firstHeader(letterhead).blocks);
		expect(textColour(header.dom, 'First Partner')).toBe(navy);
		expect(textColour(header.dom, '12 Example Street')).toBe(black);
		header.destroy();
	});

	it('is white on a dark fill and keeps a run colour in story previews', async () => {
		const { model } = await wordShapes();
		const preview = document.createElement('div');
		preview.append(renderBlocks(model.blocks, model));
		expect(textColour(preview, 'Red box')).toBe(white);
		const letterhead = (await loadDocx(await letterheadFixture())).model;
		const header = document.createElement('div');
		header.append(renderBlocks(firstHeader(letterhead).blocks, letterhead));
		expect(textColour(header, 'First Partner')).toBe(navy);
		expect(textColour(header, '12 Example Street')).toBe(black);
	});

	it('is white on a dark fill and keeps a run colour in Print Layout', async () => {
		const { model } = await wordShapes();
		const handle = renderPrintLayout(layoutDocumentModel(model, createFakeMeasurer()), undefined, {
			model,
		});
		expect(textColour(handle.element, 'Red box')).toBe(white);
		const letterhead = (await loadDocx(await letterheadFixture())).model;
		const sheet = document.createElement('div');
		decoratePages(letterhead, [firstPage], [sheet]);
		expect(textColour(sheet, 'First Partner')).toBe(navy);
		expect(textColour(sheet, '12 Example Street')).toBe(black);
	});
});
