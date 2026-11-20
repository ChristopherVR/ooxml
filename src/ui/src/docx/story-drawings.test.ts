// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { EditorState, NodeSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { loadDocx, shapeColorCss, type DocumentModel } from 'ooxml-core/docx';
import { createFakeMeasurer, layoutDocumentModel } from 'ooxml-core/docx/layout';
import { drawingsFixture } from '../../../core/docx/test-support/drawings-fixture';
import { buildHeaderElement } from './header-footer-view';
import { ImageMediaCache, imageNodeView } from './image-media';
import { modelToDoc } from './model-adapter';
import { renderPrintLayout } from './print-layout';
import { selectedTextBox, updateTextBox } from './text-box-commands';
import { DocxEditorElement, registerDocxEditor } from './index';
import { must } from './test-support';

registerDocxEditor();

const load = async () => loadDocx(await drawingsFixture());

/** An editing surface over the body with the editor's picture and shape node views. */
function surface(model: DocumentModel, cache = new ImageMediaCache(() => undefined)): EditorView {
	return new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc: modelToDoc(model) }),
		nodeViews: {
			image: imageNodeView(cache, { theme: () => model.theme, model: () => model }),
		},
	});
}
const textBox = (root: ParentNode, text: string) =>
	must(
		[...root.querySelectorAll<HTMLElement>('.dve-text-box-content')].find((element) =>
			element.textContent?.includes(text),
		),
		`text box with ${text}`,
	);
const styledGrey = (model: DocumentModel) => {
	const shape = model.blocks
		.flatMap((block) => (block.type === 'paragraph' ? block.runs : []))
		.find((run) => run.image?.shape?.fontColor)?.image?.shape;
	return must(shapeColorCss(must(shape?.fontColor, 'font colour'), model.theme), 'grey');
};
const rgb = (hex: string) =>
	`rgb(${[1, 3, 5].map((i) => Number.parseInt(hex.slice(i, i + 2), 16)).join(', ')})`;

describe('pictures in Draft view story previews', () => {
	afterEach(() => vi.restoreAllMocks());

	it('show from the body media cache, reuse its URL on every re-render and release with it', async () => {
		const { model, media } = await load();
		const created: string[] = [];
		const revoked: string[] = [];
		URL.createObjectURL = vi.fn(() => {
			created.push(`blob:logo-${created.length}`);
			return created.at(-1)!;
		});
		URL.revokeObjectURL = vi.fn((url: string) => void revoked.push(url));
		const cache = new ImageMediaCache((part) => media?.get(part));
		const urlFor = (part: string, type: string) => cache.urlFor(part, type);
		for (let render = 0; render < 3; render++) {
			const header = must(buildHeaderElement(model, 'en', 0, urlFor), 'header');
			const logo = must(header.querySelector<HTMLImageElement>('img[data-docx-image]'), 'logo');
			expect(logo.getAttribute('src')).toBe('blob:logo-0');
		}
		expect(created).toEqual(['blob:logo-0']);
		cache.release();
		expect(revoked).toEqual(['blob:logo-0']);
	});

	it('leave floating drawings out of the header while inline ones show', async () => {
		const { model, media } = await load();
		URL.createObjectURL = vi.fn(() => 'blob:logo');
		const cache = new ImageMediaCache((part) => media?.get(part));
		const header = must(
			buildHeaderElement(model, 'en', 0, (part, type) => cache.urlFor(part, type)),
			'header',
		);
		expect(header.querySelectorAll('img[data-docx-image]')).toHaveLength(1);
		expect(header.querySelector('.dve-shape')).toBeNull();
		expect(header.querySelector('.dve-text-box-content')).toBeNull();
		expect(header.textContent).toContain('Letterhead');
	});
});

describe('text boxes written with txBox="1"', () => {
	it('are drawn with their frame and paragraphs on the editing surface', async () => {
		const { model } = await load();
		const view = surface(model);
		expect(view.dom.querySelector('.dve-image-placeholder')).toBeNull();
		const floating = textBox(view.dom, 'Floating note');
		expect(
			floating.closest('.dve-shape')?.querySelector('svg path')?.getAttribute('stroke'),
		).toMatch(/^#C00000$/i);
		expect(floating.querySelector('[style*="font-weight"], strong')?.textContent).toBe(
			'Floating note',
		);
		textBox(view.dom, 'Inline box');
		view.destroy();
	});

	it('are drawn on Print Layout pages, floating and inline', async () => {
		const { model } = await load();
		const handle = renderPrintLayout(layoutDocumentModel(model, createFakeMeasurer()), undefined, {
			model,
		});
		const page = must(handle.element.querySelector<HTMLElement>('.dve-print-page'), 'page');
		expect(page.querySelector('.dve-print-picture-missing')).toBeNull();
		expect(textBox(page, 'Floating note').closest('.dve-print-float')).not.toBeNull();
		expect(textBox(page, 'Inline box').closest('.dve-print-float')).toBeNull();
	});

	it('still edit as text boxes, and show the edit as it will be saved', async () => {
		const { model } = await load();
		const view = surface(model);
		let pos = -1;
		view.state.doc.descendants((node, at) => {
			if (node.attrs.textBoxEditable) pos = at;
		});
		view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, pos)));
		const selected = must(selectedTextBox(view.state), 'selected text box');
		expect(selected.settings.lines).toEqual(['Inline box']);
		updateTextBox(view, pos, { ...selected.settings, lines: ['Edited', 'Twice'], border: false });
		const box = textBox(view.dom, 'Edited');
		expect([...box.querySelectorAll('p')].map((p) => p.textContent)).toEqual(['Edited', 'Twice']);
		expect(box.closest('.dve-shape')?.querySelector('svg path')?.getAttribute('stroke')).toBe(
			'none',
		);
		view.destroy();
	});
});

describe('shape text colour from the style font reference', () => {
	it('colours the paragraphs on the editing surface and in Print Layout alike', async () => {
		const { model } = await load();
		const grey = rgb(styledGrey(model));
		const view = surface(model);
		expect(textBox(view.dom, 'Styled text').style.color).toBe(grey);
		view.destroy();
		const handle = renderPrintLayout(layoutDocumentModel(model, createFakeMeasurer()), undefined, {
			model,
		});
		expect(textBox(handle.element, 'Styled text').style.color).toBe(grey);
	});
});

describe('pictures that arrive after they are drawn', () => {
	afterEach(() => {
		document.body.replaceChildren();
		vi.restoreAllMocks();
	});

	it('repaint in Draft view header previews when the media cache refreshes', async () => {
		const { model, media } = await load();
		let next = 0;
		URL.createObjectURL = vi.fn(() => `blob:late-${next++}`);
		URL.revokeObjectURL = vi.fn();
		const late = new Map<string, Uint8Array>();
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		editor.setLoadedDocument({ model, media: late, save: async () => new Uint8Array() });
		document.body.append(editor);
		const logo = () =>
			editor.shadowRoot!.querySelector<HTMLImageElement>('.dve-header img[data-docx-image]');
		expect(logo()?.getAttribute('src')).toBeNull();
		late.set('word/media/logo.png', must(media?.get('word/media/logo.png'), 'logo bytes'));
		(editor as unknown as { core: { imageMedia: ImageMediaCache } }).core.imageMedia.refresh();
		expect(logo()?.getAttribute('src')).toMatch(/^blob:late-/);
	});

	it('repaint inside a text box on the editing surface when the media cache refreshes', async () => {
		const { model, media } = await load();
		let next = 0;
		URL.createObjectURL = vi.fn(() => `blob:box-${next++}`);
		URL.revokeObjectURL = vi.fn();
		const box = must(
			model.blocks
				.flatMap((block) => (block.type === 'paragraph' ? block.runs : []))
				.find((run) => run.image?.shape?.fontColor)?.image?.textBoxContent,
			'styled shape content',
		);
		const paragraph = box.blocks[0]!;
		if (paragraph.type === 'paragraph')
			paragraph.runs.push({
				text: '',
				image: {
					relId: 'rId9',
					partName: 'word/media/logo.png',
					contentType: 'image/png',
					widthPx: 16,
					heightPx: 8,
				},
			});
		const late = new Map<string, Uint8Array>();
		const cache = new ImageMediaCache((part) => late.get(part));
		const view = surface(model, cache);
		const picture = () => textBox(view.dom, 'Styled text').querySelector('img[data-docx-image]');
		expect(picture()?.getAttribute('src')).toBeNull();
		late.set('word/media/logo.png', must(media?.get('word/media/logo.png'), 'logo bytes'));
		cache.refresh();
		expect(picture()?.getAttribute('src')).toMatch(/^blob:box-/);
		view.destroy();
	});
});
