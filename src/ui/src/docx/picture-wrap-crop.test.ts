// @vitest-environment jsdom
import { readFile } from 'node:fs/promises';
import { loadDocx } from 'ooxml-core/docx';
import { afterEach, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { EditorState, NodeSelection } from 'prosemirror-state';
import { undo } from 'prosemirror-history';
import { EditorView } from 'prosemirror-view';
import { DocxEditorElement, registerDocxEditor } from './index';
import { at, paragraphAt } from './test-support';
import { computeMenuItems } from './context-menu-items';
import { buildPictureFormatPanel, syncPictureFormat } from './picture-format-ribbon';
import { strings } from './localization-strings';
import { isPictureNode, setPictureCropAt, setPictureWrap } from './picture-format-commands';
import { createPictureDialog } from './picture-dialog';
import { docToModel, modelToDoc } from './model-adapter';
import { defineRibbon } from '../ribbon/ribbon-tabs';
import { schema } from './schema';

registerDocxEditor();

const NS = [
	'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
	'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
	'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
	'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
	'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
].join(' ');

async function pictureDocx(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document ${NS}><w:body><w:p><w:r><w:t>Logo </w:t></w:r><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="952500" cy="476250"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="7" name="Picture 1"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="7" name=""/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>',
	);
	zip.file('word/media/image1.png', new Uint8Array([137, 80, 78, 71]));
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/></Types>',
	);
	return zip.generateAsync({ type: 'uint8array' });
}

async function openEditor() {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	document.body.append(editor);
	await editor.load(await pictureDocx());
	const view = (editor as unknown as { view: EditorView }).view;
	let pos = -1;
	view.state.doc.descendants((node, at) => {
		if (node.type.name === 'image') pos = at;
	});
	const select = () =>
		view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, pos)));
	const image = () => at(paragraphAt(editor.documentModel!.blocks, 0).runs, 1).image!;
	const xml = async () =>
		(await JSZip.loadAsync(await editor.saveBytes())).file('word/document.xml')!.async('string');
	return { editor, view, root: editor.shadowRoot!, select, image, xml };
}

describe('picture Wrap Text and Crop', () => {
	afterEach(() => document.body.replaceChildren());

	it('shows Picture Format for a selected picture and wraps it as one undoable step', async () => {
		const { root, view, select, image, xml } = await openEditor();
		const tab = root.querySelector<HTMLElement>('[data-ribbon-tab="picture-format"]')!;
		expect(tab.hasAttribute('data-tab-hidden')).toBe(true);
		select();
		expect(tab.hasAttribute('data-tab-hidden')).toBe(false);
		const wrap = tab.querySelector<HTMLSelectElement>('select[aria-label="Wrap text"]')!;
		expect(wrap.value).toBe('inline');
		wrap.value = 'square';
		wrap.dispatchEvent(new Event('change'));
		expect(image()).toMatchObject({ anchored: true, placement: { wrap: 'square' } });
		expect(view.state.selection).toBeInstanceOf(NodeSelection);
		expect(await xml()).toContain('<wp:wrapSquare wrapText="bothSides"/>');
		wrap.value = 'behind';
		wrap.dispatchEvent(new Event('change'));
		expect(image().placement).toMatchObject({ wrap: 'none', behindText: true });
		undo(view.state, view.dispatch);
		expect(image().placement?.wrap).toBe('square');
		undo(view.state, view.dispatch);
		expect(image().anchored).toBeUndefined();
		expect(await xml()).toContain('<wp:inline distT="0"');
	});

	it('crops in the Format Picture dialog, following the size, and undoes it', async () => {
		const { root, view, select, image, xml } = await openEditor();
		select();
		root
			.querySelector<HTMLButtonElement>('[data-ribbon-tab="picture-format"] [aria-label="Crop"]')!
			.click();
		const dialog = root.querySelector<HTMLElement>('.dve-picture-dialog')!;
		const field = (label: string) =>
			dialog.querySelector<HTMLInputElement>(`.dve-picture-crop [aria-label="${label}"]`)!;
		expect(dialog.hidden).toBe(false);
		expect(field('Left').value).toBe('0');
		field('Left').value = '10';
		field('Left').dispatchEvent(new Event('input'));
		field('Bottom').value = '50';
		field('Bottom').dispatchEvent(new Event('input'));
		expect(dialog.querySelector<HTMLInputElement>('[aria-label="Width"]')!.value).toBe('90');
		expect(dialog.querySelector<HTMLInputElement>('[aria-label="Height"]')!.value).toBe('25');
		[...dialog.querySelectorAll('button')].find((button) => button.textContent === 'OK')!.click();
		expect(image()).toMatchObject({
			widthPx: 90,
			heightPx: 25,
			pictureFill: { crop: { cropLeft: 0.1, cropBottom: 0.5 } },
		});
		expect(await xml()).toContain('<a:blip r:embed="rId1"/><a:srcRect l="10000" b="50000"/>');
		undo(view.state, view.dispatch);
		expect(image()).toMatchObject({ widthPx: 100, heightPx: 50 });
		expect(image().pictureFill?.crop).toBeUndefined();
	});

	it('saves a picture position, transform and alt text from Picture Format', async () => {
		const { root, view, select, image, xml } = await openEditor();
		select();
		const panel = root.querySelector<HTMLElement>('[data-ribbon-tab="picture-format"]')!;
		const choose = (label: string, value: string) => {
			const control = panel.querySelector<HTMLSelectElement>(`select[aria-label="${label}"]`)!;
			control.value = value;
			control.dispatchEvent(new Event('change'));
		};
		const mode = root.querySelector<HTMLSelectElement>('select[aria-label="Layout view"]')!;
		mode.value = 'print';
		mode.dispatchEvent(new Event('change'));
		choose('Position', 'bottom-right');
		// A presence-like metadata transaction does not invalidate the pending ribbon command.
		view.dispatch(view.state.tr.setMeta('test-presence', true));
		await vi.waitFor(() =>
			expect(image().placement).toMatchObject({ align: 'right', alignV: 'bottom' }),
		);
		choose('Rotate', 'right');
		await vi.waitFor(() => expect(image().transform).toEqual({ rotation: 90 }));
		panel.querySelector<HTMLButtonElement>('[aria-label="Alt Text"]')!.click();
		await vi.waitFor(() =>
			expect(root.querySelector('.dve-arrange-dialog textarea')).not.toBeNull(),
		);
		root.querySelector<HTMLTextAreaElement>('.dve-arrange-dialog textarea')!.value =
			'A picture caption';
		const ok = [...root.querySelectorAll<HTMLButtonElement>('.dve-arrange-dialog button')].find(
			(button) => button.textContent === 'OK',
		)!;
		ok.click();
		expect(image().altText).toBe('A picture caption');
		expect(await xml()).toMatch(/descr="A picture caption".*<a:xfrm rot="5400000"/);
	});

	it('refuses a crop that leaves nothing showing', async () => {
		const { root, select } = await openEditor();
		select();
		root
			.querySelector<HTMLButtonElement>('[data-ribbon-tab="picture-format"] [aria-label="Crop"]')!
			.click();
		const dialog = root.querySelector<HTMLElement>('.dve-picture-dialog')!;
		const left = dialog.querySelector<HTMLInputElement>('.dve-picture-crop [aria-label="Left"]')!;
		const right = dialog.querySelector<HTMLInputElement>('.dve-picture-crop [aria-label="Right"]')!;
		left.value = '60';
		right.value = '40';
		right.dispatchEvent(new Event('input'));
		const ok = [...dialog.querySelectorAll('button')].find(
			(button) => button.textContent === 'OK',
		)!;
		expect(ok.disabled).toBe(true);
		expect(dialog.querySelector('[role="alert"]')!.textContent).toBe(
			'The crop must leave part of the picture showing.',
		);
	});

	it('offers Wrap Text in the context menu of a selected picture', () => {
		const base = {
			readOnly: false,
			hasSelection: true,
			onLink: false,
			inTable: false,
			canReadClipboard: true,
			canTable: () => false,
		};
		expect(computeMenuItems(base).some((item) => item.id.startsWith('picture-'))).toBe(false);
		const items = computeMenuItems({ ...base, pictureWrap: 'tight' });
		const wraps = items.filter((item) => item.id.startsWith('picture-wrap-'));
		expect(wraps.map((item) => item.id)).toEqual([
			'picture-wrap-inline',
			'picture-wrap-square',
			'picture-wrap-tight',
			'picture-wrap-through',
			'picture-wrap-topAndBottom',
			'picture-wrap-behind',
			'picture-wrap-front',
		]);
		expect(wraps[0]).toMatchObject({ heading: 'Wrap text', separatorBefore: true });
		expect(wraps.filter((item) => item.checked).map((item) => item.id)).toEqual([
			'picture-wrap-tight',
		]);
		expect(wraps[2]!.command).toEqual({
			kind: 'ribbon',
			action: { type: 'picture', key: 'wrap', value: 'tight' },
		});
		expect(items.find((item) => item.id === 'picture-format')).toBeDefined();
	});

	it('offers Picture Format for VML floats even without a modeled picture fill', () => {
		const image = (attrs: Record<string, unknown>) =>
			schema.nodes.image!.create({ partName: 'word/media/a.png', ...attrs });
		expect(isPictureNode(image({ relId: 'rId1' }))).toBe(false);
		expect(isPictureNode(image({ relId: 'rId1', vmlFloat: true, anchored: true }))).toBe(true);
		expect(isPictureNode(image({ relId: 'rId1', pictureFill: '{"kind":"picture"}' }))).toBe(true);
		expect(isPictureNode(image({ relId: '' }))).toBe(true);
	});

	it('allows wrapping a VML float and explains why Crop is disabled', async () => {
		const path =
			'../core/docx/layout/fixtures/vml-floats/issue-345-vml-absolute-picture-offset.docx';
		const loaded = await loadDocx(await readFile(path));
		const model = structuredClone(loaded.model);
		const header = model.sections?.[0]?.headers?.default;
		if (!header) throw new Error('Expected a default header');
		const story = { ...model, blocks: header.blocks };
		const doc = modelToDoc(story);
		let pos = -1;
		doc.descendants((node, at) => {
			if (node.type.name === 'image') pos = at;
		});
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({ doc, selection: NodeSelection.create(doc, pos) }),
		});
		try {
			defineRibbon();
			const toolbar = document.createElement('office-ui-ribbon');
			toolbar.innerHTML = '<div data-ribbon-tab="picture-format"></div>';
			buildPictureFormatPanel(toolbar.firstElementChild as HTMLElement);
			syncPictureFormat(toolbar, view.state, true);
			const crop = toolbar.querySelector<HTMLButtonElement>('[aria-label="Crop"]');
			expect(crop?.disabled).toBe(true);
			expect(crop?.getAttribute('aria-description')).toMatch(/VML/i);
			const dialog = createPictureDialog(() => view);
			dialog.open(pos);
			expect(dialog.element.querySelector<HTMLFieldSetElement>('.dve-picture-crop')?.disabled).toBe(
				true,
			);
			expect(dialog.element.textContent).toContain(crop?.getAttribute('aria-description'));
			expect(
				setPictureCropAt(view, pos, { cropLeft: 0.1, cropTop: 0, cropRight: 0, cropBottom: 0 }),
			).toBe(false);
			const width = dialog.element.querySelector<HTMLInputElement>('[aria-label="Width"]');
			if (!width) throw new Error('Expected a width field');
			width.value = '320';
			width.dispatchEvent(new Event('input'));
			dialog.element.querySelector<HTMLButtonElement>('.dve-dialog-primary')?.click();
			expect(setPictureWrap(view, 'square', pos)).toBe(true);
			expect(setPictureWrap(view, 'inline', pos)).toBe(false);
			header.blocks = docToModel(view.state.doc, story).blocks;
			const saved = await loadDocx(await loaded.save(model));
			const images = saved.model.sections?.[0]?.headers?.default?.blocks.flatMap((block) =>
				block.type === 'paragraph' ? block.runs.flatMap((run) => run.image ?? []) : [],
			);
			expect(images).toEqual([
				expect.objectContaining({
					vmlFloat: true,
					widthPx: 320,
					placement: expect.objectContaining({ wrap: 'square' }),
				}),
			]);
		} finally {
			view.destroy();
		}
	});

	it('translates every label of the Picture Format tab', () => {
		const panel = document.createElement('div');
		buildPictureFormatPanel(panel);
		for (const element of panel.querySelectorAll<HTMLElement>(
			'[aria-label], [data-label], option',
		)) {
			const label = (
				element.getAttribute('aria-label') ??
				element.dataset.label ??
				element.textContent ??
				''
			).replace(/ controls$/, '');
			if (label) expect(label in strings.en, label).toBe(true);
		}
	});
});
