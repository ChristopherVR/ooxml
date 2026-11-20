// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
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

/** What a text box and a line in a story must show: the box's paragraphs and two drawn shapes. */
function expectDrawnLetterhead(root: HTMLElement) {
	expect(root.querySelector('.dve-print-picture-missing')).toBeNull();
	expect(root.querySelector('.dve-image-placeholder')).toBeNull();
	expect(root.textContent).not.toContain('Shape');
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
	const strokes = [...root.querySelectorAll('svg path')].map((path) => path.getAttribute('stroke'));
	expect(strokes).toContain('#2E75B6');
}

describe('text boxes and lines in headers', () => {
	it('draws them with their paragraphs in the draft header story', async () => {
		const { model } = await loadDocx(await letterheadFixture());
		const header = must(buildHeaderElement(model, 'en'), 'header story');
		expectDrawnLetterhead(header);
		expect(header.querySelectorAll('.dve-shape svg')).toHaveLength(2);
	});

	it('draws them on the Print Layout page instead of empty placeholder boxes', async () => {
		const { model } = await loadDocx(await letterheadFixture());
		const sheet = document.createElement('div');
		decoratePages(model, [firstPage], [sheet]);
		expectDrawnLetterhead(sheet);
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
		expect(must(floats[0]?.querySelector('.dve-text-box-content'), 'box').textContent).toContain(
			'First Partner',
		);
	});

	it('draws the paragraphs over the shape on the editing surface', async () => {
		const { model } = await loadDocx(await letterheadFixture());
		const doc = modelToDoc({ ...model, blocks: firstHeader(model).blocks });
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({ doc }),
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
