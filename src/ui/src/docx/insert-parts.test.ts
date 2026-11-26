// @vitest-environment jsdom
import {
	createDocument,
	loadDocx,
	saveDocx,
	type DocumentModel,
	type GlossaryEntry,
	type Paragraph,
} from 'ooxml-core/docx';
import JSZip from 'jszip';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { beforeAll, describe, expect, it } from 'vitest';
import * as headerFooter from './header-footer-commands';
import * as sections from './section-commands';
import * as designs from './header-footer-design-commands';
import { createRibbon } from './ribbon';
import { docToModel, modelToDoc } from './model-adapter';
import { mountInsertParts, quickPartsMenuAction, syncInsertParts } from './insert-parts-ribbon';
import {
	insertDocumentProperty,
	insertGlossaryEntry,
	selectionBlocks,
} from './quick-parts-commands';
import { insertSignatureLine, insertWebVideo } from './insert-parts-insert';
import { schema } from './schema';
import { EDITOR_LOCALES, translate, type LocalizationKey } from './localization';

beforeAll(() => {
	const rects = { length: 0, item: () => null, [Symbol.iterator]: function* () {} };
	Object.assign(Range.prototype, {
		getClientRects: () => rects,
		getBoundingClientRect: () => ({ top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0 }),
	});
});

const tools = { ...headerFooter, ...sections };
const newHeaderFooterId = headerFooter.newHeaderFooterId;
const withHeaderFooterDesign = (
	...args: Parameters<typeof designs.withHeaderFooterDesign> extends [unknown, ...infer R]
		? R
		: never
) => designs.withHeaderFooterDesign(tools, ...args);
const withoutHeaderFooter = (
	...args: Parameters<typeof designs.withoutHeaderFooter> extends [unknown, ...infer R] ? R : never
) => designs.withoutHeaderFooter(tools, ...args);

function modelWith(text: string): DocumentModel {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p1', runs: [{ text }] }];
	model.properties = { title: 'Annual report', creator: 'Ada' };
	return model;
}
function viewOf(model: DocumentModel) {
	return new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc: modelToDoc(model), schema }),
	});
}

describe('header and footer designs', () => {
	it('writes a design with fields and tab stops, and reads it back after a save', async () => {
		const model = modelWith('Body');
		const next = withHeaderFooterDesign(model, 'headers', 'titleAndAuthor', newHeaderFooterId);
		const header = next.sections![0]!.headers!.default!;
		expect(header.blocks).toHaveLength(1);
		const paragraph = header.blocks[0] as Paragraph;
		expect(paragraph.tabStops?.map((stop) => stop.align)).toEqual(['center', 'right']);
		expect(paragraph.runs.map((run) => run.contentControl?.alias)).toEqual([
			'Title',
			undefined,
			undefined,
			'Author',
		]);
		expect(paragraph.runs.map((run) => run.text)).toEqual(['Annual report', '\t', '\t', 'Ada']);
		const bytes = await saveDocx(next);
		const zip = await JSZip.loadAsync(bytes);
		const part = await zip.file(header.partName!)!.async('string');
		expect(part).toContain('<w:alias w:val="Title"/>');
		expect(part).toContain('<w:tab w:val="right"');
		const again = await loadDocx(bytes);
		expect(again.model.sections![0]!.headers!.default!.blocks).toHaveLength(1);
	});

	it('replaces an existing story with another design', () => {
		const model = withHeaderFooterDesign(
			modelWith('Body'),
			'footers',
			'dateAndPage',
			newHeaderFooterId,
		);
		const centered = withHeaderFooterDesign(model, 'footers', 'centeredTitle', newHeaderFooterId);
		const paragraph = centered.sections![0]!.footers!.default!.blocks[0] as Paragraph;
		expect(paragraph.align).toBe('center');
		expect(centered.sections![0]!.footers!.default!.partName).toBe(
			model.sections![0]!.footers!.default!.partName,
		);
	});

	it('removes a header so the saved section no longer references it', async () => {
		const model = withHeaderFooterDesign(modelWith('Body'), 'headers', 'title', newHeaderFooterId);
		const first = await loadDocx(await saveDocx(model));
		const removed = withoutHeaderFooter(first.model, 'headers');
		expect(removed.sections![0]!.headers).toBeUndefined();
		const saved = await JSZip.loadAsync(await first.save(removed));
		expect(await saved.file('word/document.xml')!.async('string')).not.toContain('headerReference');
	});

	it('adds the gallery designs to the Header and Footer menus and maps them to actions', async () => {
		const ribbon = createRibbon();
		await mountInsertParts(ribbon);
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		const choose = (label: string, value: string) => {
			const select = ribbon.querySelector<HTMLSelectElement>(`select[aria-label="${label}"]`)!;
			expect([...select.options].map((option) => option.value)).toContain('threeColumns');
			select.value = value;
			select.dispatchEvent(new Event('change'));
		};
		choose('Header', 'edit');
		choose('Footer', 'remove');
		choose('Footer', 'dateAndPage');
		choose('Header', 'even');
		expect(seen).toEqual([
			{ type: 'headerFooterEdit', kind: 'header' },
			{ type: 'headerFooterRemove', kind: 'footer' },
			{ type: 'headerFooter', kind: 'footer', slot: 'default', design: 'dateAndPage' },
			{ type: 'headerFooter', kind: 'header', slot: 'even' },
		]);
	});
});

describe('quick parts', () => {
	it('inserts a document property control bound to the property, with its value', async () => {
		const model = modelWith('Hello');
		const view = viewOf(model);
		view.dispatch(view.state.tr.setSelection(TextSelection.atEnd(view.state.doc)));
		expect(insertDocumentProperty(view, model, 'title', 'fr')).toBe(true);
		const back = docToModel(view.state.doc, model);
		const run = (back.blocks[0] as Paragraph).runs.find((item) => item.contentControl);
		expect(run?.text).toBe('Annual report');
		expect(run?.contentControl).toMatchObject({ dataBound: true, alias: 'Titre' });
		const saved = await JSZip.loadAsync(await saveDocx(back));
		const xml = await saved.file('word/document.xml')!.async('string');
		expect(xml).toContain('w:storeItemID="{6C3C8BC8-F283-45AE-878A-BAB7291924A1}"');
	});

	it('shows the bracketed name while the property is empty', () => {
		const model = modelWith('x');
		const view = viewOf(model);
		insertDocumentProperty(view, model, 'company', 'en');
		expect(view.state.doc.textContent).toContain('[Company]');
	});

	it('saves a selection as blocks and inserts a gallery entry back', () => {
		const model = modelWith('Dear customer');
		const view = viewOf(model);
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 6)));
		const blocks = selectionBlocks(view, model);
		const first = blocks?.[0] as Paragraph | undefined;
		expect(first?.runs.map((run) => run.text).join('')).toBe('Dear ');
		const entry: GlossaryEntry = {
			name: 'Greeting',
			gallery: 'docParts',
			category: 'General',
			blocks: blocks!,
		};
		view.dispatch(view.state.tr.setSelection(TextSelection.atEnd(view.state.doc)));
		expect(insertGlossaryEntry(view, model, entry)).toBe(true);
		expect(view.state.doc.childCount).toBe(1);
		expect(view.state.doc.textContent).toBe('Dear customerDear ');
		const multi: GlossaryEntry = {
			...entry,
			blocks: [
				{ type: 'paragraph', id: 'a', runs: [{ text: 'One' }] },
				{ type: 'paragraph', id: 'b', runs: [{ text: 'Two' }] },
			],
		};
		insertGlossaryEntry(view, model, multi);
		expect(view.state.doc.childCount).toBeGreaterThan(1);
	});

	it('lists Quick Parts entries in the menu and shows Online Video only for an allow-list', async () => {
		const ribbon = createRibbon();
		await mountInsertParts(ribbon);
		syncInsertParts(ribbon, [], []);
		const group = ribbon.querySelector<HTMLElement>('[data-online-video]')!;
		expect(group.hidden).toBe(true);
		const entry: GlossaryEntry = {
			name: 'Sign-off',
			gallery: 'docParts',
			category: 'General',
			blocks: [],
		};
		syncInsertParts(ribbon, [entry], ['video.example.com']);
		expect(group.hidden).toBe(false);
		const select = ribbon.querySelector<HTMLSelectElement>('select[aria-label="Quick Parts"]')!;
		expect(select.options[0]!.textContent).toBe('Sign-off');
		expect(select.options[0]!.value).toMatch(/^entry:/);
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		select.value = select.options[0]!.value;
		select.dispatchEvent(new Event('change'));
		expect(seen[0]).toMatchObject({ type: 'quickPart', command: { kind: 'entry' } });
		expect(quickPartsMenuAction('organizer')).toEqual({
			type: 'insertParts',
			kind: 'buildingBlocks',
		});
		expect(quickPartsMenuAction('property:company')).toMatchObject({
			command: { kind: 'property', id: 'company' },
		});
	});
});

describe('signature line and online video insertion', () => {
	it('inserts a signature line picture that survives the model round trip', async () => {
		const model = modelWith('Sign below');
		const view = viewOf(model);
		const staged = new Map<string, Uint8Array>();
		await insertSignatureLine(
			view,
			{ id: '{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}', allowComments: true, showSignDate: true },
			(bytes, contentType) => {
				const name = `word/media/staged${staged.size}.png`;
				staged.set(name, bytes);
				expect(contentType).toBe('image/png');
				return name;
			},
		);
		const back = docToModel(view.state.doc, model);
		const image = (back.blocks[0] as Paragraph).runs.find((run) => run.image)?.image;
		expect(image?.signatureLine?.id).toBe('{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}');
		const bytes = await saveDocx(
			back,
			new Map([...staged].map(([name, data]) => [name, { bytes: data, contentType: 'image/png' }])),
		);
		const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
		expect(xml).toContain('issignatureline="t"');
	});

	it('inserts an online video poster carrying the embed', () => {
		const model = modelWith('Watch');
		const view = viewOf(model);
		insertWebVideo(
			view,
			{
				url: 'https://video.example.com/embed/1',
				embeddedHtml: '<iframe src="https://video.example.com/embed/1"></iframe>',
				width: 560,
				height: 315,
			},
			() => 'word/media/poster.png',
		);
		const back = docToModel(view.state.doc, model);
		const image = (back.blocks[0] as Paragraph).runs.find((run) => run.image)?.image;
		expect(image?.webVideo?.url).toBe('https://video.example.com/embed/1');
	});
});

describe('locales', () => {
	it('translates the new menu and dialog strings in every locale', async () => {
		await import('./locales/insert-parts-translations');
		const keys = [
			'Quick Parts',
			'Signature Setup',
			'Building Blocks Organizer',
			'Edit Header',
			'Insert Video',
		] as unknown as LocalizationKey[];
		for (const locale of EDITOR_LOCALES)
			if (locale !== 'en') for (const key of keys) expect(translate(locale, key)).not.toBe(key);
	});
});
