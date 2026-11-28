// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import {
	createDocument,
	loadDocx,
	type DocumentModel,
	type Paragraph,
	type PictureWatermarkSpec,
	type WatermarkSpec,
} from 'ooxml-core/docx';
import { createWatermarkDialog, type WatermarkHost } from './watermark-dialog';
import { currentWatermark, withWatermark, type AnyWatermark } from './watermark-commands';
import { watermarkElement } from './print-header-footer';
import { WASHOUT_FILTER } from './print-picture-watermark';
import { WASHOUT_CSS_FILTER } from 'ooxml-core/docx';
import { DocxEditorElement, registerDocxEditor } from './index';

registerDocxEditor();

const PNG = Uint8Array.from(
	atob(
		'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
	),
	(char) => char.charCodeAt(0),
);
const picture: PictureWatermarkSpec = {
	kind: 'picture',
	partName: 'word/media/seal.png',
	contentType: 'image/png',
	widthPt: 300,
	heightPt: 150,
	scale: 'auto',
	washout: true,
	title: 'seal',
};
const text: WatermarkSpec = {
	text: 'DRAFT',
	color: '#c0c0c0',
	semitransparent: true,
	layout: 'diagonal',
};
let counter = 0;
const newId = () => `w${++counter}`;
const headerRuns = (model: DocumentModel) =>
	(model.sections![0]!.headers!.default!.blocks[0] as Paragraph).runs;

describe('picture watermark commands', () => {
	it('adds a picture watermark, replaces a text one with it and back, and removes it', () => {
		const withText = withWatermark(createDocument(), text, newId);
		const withPicture = withWatermark(withText, picture, newId);
		const runs = headerRuns(withPicture).filter((run) => run.image);
		expect(runs).toHaveLength(1);
		expect(runs[0]!.image).toMatchObject({
			partName: 'word/media/seal.png',
			unsupported: 'Watermark',
			pictureWatermark: picture,
		});
		expect(runs[0]!.image!.watermark).toBeUndefined();
		expect(currentWatermark(withPicture)).toEqual(picture);
		const back = withWatermark(withPicture, text, newId);
		expect(headerRuns(back).filter((run) => run.image?.pictureWatermark)).toHaveLength(0);
		expect(currentWatermark(back)).toEqual(text);
		expect(currentWatermark(withWatermark(withPicture, undefined, newId))).toBeUndefined();
	});
});

describe('picture watermark layer', () => {
	const content = (spec: PictureWatermarkSpec) => ({
		blocks: [
			{
				type: 'paragraph' as const,
				id: 'p',
				runs: [
					{
						text: '',
						image: {
							relId: '',
							partName: spec.partName,
							contentType: spec.contentType,
							widthPx: 0,
							heightPx: 0,
							unsupported: 'Watermark',
							pictureWatermark: spec,
						},
					},
				],
			},
		],
	});
	const page = {
		widthPx: 816,
		heightPx: 1056,
		marginLeftPx: 96,
		marginRightPx: 96,
		marginTopPx: 96,
		marginBottomPx: 96,
	};

	it('centres the picture on the margins behind the text, washed out', () => {
		const layer = watermarkElement(content(picture), page, 'all', (part) => `blob:${part}`)!;
		expect(layer.tagName).toBe('IMG');
		expect((layer as HTMLImageElement).src).toBe('blob:word/media/seal.png');
		expect(layer.classList.contains('dve-print-watermark')).toBe(true);
		expect(layer.style.width).toBe('400px');
		expect(layer.style.height).toBe('200px');
		expect(layer.style.left).toBe('208px');
		expect(layer.style.top).toBe('428px');
		expect(layer.style.filter).toBe(WASHOUT_FILTER);
		expect(layer.style.pointerEvents).toBe('none');
		const plain = watermarkElement(content({ ...picture, washout: false }), page)!;
		expect(plain.style.filter).toBe('');
		expect(plain.classList.contains('dve-print-picture-missing')).toBe(true);
	});
});

describe('Printed Watermark dialog, picture', () => {
	afterEach(() => document.body.replaceChildren());

	const open = (overrides: Partial<WatermarkHost>, applied: Array<AnyWatermark | undefined>) => {
		const dialog = createWatermarkDialog({
			current: () => undefined,
			canEdit: () => true,
			apply: (value) => applied.push(value),
			restoreFocus: () => {},
			textArea: () => ({ widthPt: 468, heightPt: 648 }),
			...overrides,
		});
		document.body.append(dialog.element);
		dialog.open();
		const field = (name: string) =>
			dialog.element.querySelector<HTMLInputElement & HTMLSelectElement>(`[aria-label="${name}"]`)!;
		const button = (label: string) =>
			[...dialog.element.querySelectorAll('button')].find((b) => b.textContent === label)!;
		return { dialog, field, button };
	};
	const flush = () => new Promise((resolve) => setTimeout(resolve, 0));

	it('stages a chosen picture and applies its scale and washout', async () => {
		const applied: Array<AnyWatermark | undefined> = [];
		const { dialog, field, button } = open(
			{
				stagePicture: async (file) => ({
					partName: 'word/media/new.png',
					contentType: file.type,
					widthPt: 150,
					heightPt: 75,
					title: 'logo',
				}),
			},
			applied,
		);
		field('Setting').value = 'picture';
		field('Setting').dispatchEvent(new Event('change'));
		expect(button('OK').disabled).toBe(true);
		expect(dialog.element.querySelector('[role="alert"]')!.textContent).toBe(
			'Select a picture for the watermark.',
		);
		const input = dialog.element.querySelector<HTMLInputElement>('input[type="file"]')!;
		Object.defineProperty(input, 'files', {
			value: [new File([PNG], 'logo.png', { type: 'image/png' })],
			configurable: true,
		});
		input.dispatchEvent(new Event('change'));
		await flush();
		expect(button('OK').disabled).toBe(false);
		field('Scale').value = '200';
		field('Washout').checked = false;
		button('OK').click();
		await flush();
		expect(applied[0]).toEqual({
			kind: 'picture',
			partName: 'word/media/new.png',
			contentType: 'image/png',
			widthPt: 300,
			heightPt: 150,
			scale: 200,
			washout: false,
			title: 'logo',
		});
	});

	it('reopens on the current picture and keeps its size unless the scale changes', async () => {
		const loaded: PictureWatermarkSpec = { ...picture, scale: undefined as never };
		delete (loaded as { scale?: unknown }).scale;
		const applied: Array<AnyWatermark | undefined> = [];
		const { field, button } = open(
			{
				current: () => loaded,
				pictureSize: async () => ({ widthPt: 100, heightPt: 50 }),
			},
			applied,
		);
		expect(field('Setting').value).toBe('picture');
		expect(field('Scale').value).toBe('auto');
		expect(field('Washout').checked).toBe(true);
		button('OK').click();
		await flush();
		expect(applied[0]).toMatchObject({ widthPt: 300, heightPt: 150, scale: 'auto' });
		const again = open(
			{ current: () => loaded, pictureSize: async () => ({ widthPt: 100, heightPt: 50 }) },
			applied,
		);
		again.field('Scale').value = '50';
		again.button('OK').click();
		await flush();
		expect(applied[1]).toMatchObject({ widthPt: 50, heightPt: 25, scale: 50 });
	});
});

describe('picture watermark in the editor', () => {
	afterEach(() => document.body.replaceChildren());

	it('inserts a picture watermark from the dialog, saves its picture and reads it back', async () => {
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		editor.documentModel = createDocument();
		document.body.append(editor);
		const root = editor.shadowRoot!;
		root.querySelector<HTMLButtonElement>('.dve-ribbon [aria-label="Watermark"]')!.click();
		const dialog = root.querySelector<HTMLElement>('.dve-watermark-dialog')!;
		const setting = dialog.querySelector<HTMLSelectElement>('[aria-label="Setting"]')!;
		setting.value = 'picture';
		setting.dispatchEvent(new Event('change'));
		const input = dialog.querySelector<HTMLInputElement>('input[type="file"]')!;
		Object.defineProperty(input, 'files', {
			value: [new File([PNG], 'seal.png', { type: 'image/png' })],
			configurable: true,
		});
		input.dispatchEvent(new Event('change'));
		await new Promise((resolve) => setTimeout(resolve, 20));
		[...dialog.querySelectorAll('button')].find((b) => b.textContent === 'OK')!.click();
		await new Promise((resolve) => setTimeout(resolve, 20));
		const spec = currentWatermark(editor.documentModel!) as PictureWatermarkSpec;
		expect(spec).toMatchObject({ kind: 'picture', title: 'seal', washout: true, scale: 'auto' });
		const bytes = await editor.saveBytes();
		const zip = await JSZip.loadAsync(bytes);
		const header = Object.keys(zip.files).find((name) => /^word\/header\d*\.xml$/.test(name))!;
		const xml = await zip.file(header)!.async('string');
		expect(xml).toMatch(/id="WordPictureWatermark\d+"/);
		expect(xml).toContain('gain="19661f" blacklevel="22938f"');
		expect(await zip.file(spec.partName)!.async('uint8array')).toEqual(PNG);
		const reopened = await loadDocx(bytes);
		expect(currentWatermark(reopened.model)).toMatchObject({
			kind: 'picture',
			partName: spec.partName,
			washout: true,
			title: 'seal',
		});
	});
});

describe('washout filter', () => {
	it('maps white to white and black to light grey, as Word does', () => {
		expect(WASHOUT_FILTER).toBe(WASHOUT_CSS_FILTER);
		const [contrast, brightness] = [...WASHOUT_FILTER.matchAll(/\(([\d.]+)\)/g)].map((m) =>
			Number(m[1]),
		);
		const channel = (x: number) => Math.min(1, ((x - 0.5) * contrast! + 0.5) * brightness!);
		expect(channel(1)).toBe(1);
		expect(channel(0)).toBeCloseTo(0.8, 2);
	});
});
