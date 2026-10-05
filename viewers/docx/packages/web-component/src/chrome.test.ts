// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createDocument } from 'docx-core';
import { DocxEditorElement, registerDocxEditor } from './index';

registerDocxEditor();

function mount(): DocxEditorElement {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	const model = createDocument();
	model.blocks[0] = { type: 'paragraph', id: 'p1', runs: [{ text: 'Hello chrome' }] };
	model.warnings = ['Pictures are shown as placeholders.'];
	editor.documentModel = model;
	editor.fileName = 'Report.docx';
	document.body.append(editor);
	return editor;
}
const root = (editor: DocxEditorElement) => editor.shadowRoot!;
type Updating = HTMLElement & { updateComplete: Promise<unknown> };
/** The shadow root of a shared element inside the editor, once it has rendered. */
const inner = async (editor: DocxEditorElement, tag: string) => {
	const element = root(editor).querySelector<Updating>(tag)!;
	await element.updateComplete;
	return element.shadowRoot!;
};

describe('Word window chrome', () => {
	it('File Options changes live view and spelling preferences without editing the document', async () => {
		const editor = mount();
		const r = root(editor);
		r.querySelector<HTMLButtonElement>('.dve-file-tab')!.click();
		(await inner(editor, 'office-ui-backstage'))
			.querySelector<HTMLButtonElement>('[data-backstage-item="options"]')!
			.click();
		const toggle = async (label: string) => {
			r.querySelector<HTMLInputElement>(
				`.dve-backstage-content input[aria-label="${label}"]`,
			)!.click();
			await Promise.resolve();
		};
		await toggle('Ruler');
		expect(r.querySelector('.dve-ruler')).not.toBeNull();
		await toggle('Show paragraph marks');
		expect(r.querySelector('.dve-paper')?.hasAttribute('data-show-marks')).toBe(true);
		await toggle('Page thumbnails');
		expect(editor.hasAttribute('show-thumbnails')).toBe(true);
		await toggle('Spelling');
		expect(
			(editor as unknown as { view: import('prosemirror-view').EditorView }).view.dom.spellcheck,
		).toBe(false);
		expect(editor.dirty).toBe(false);
		await toggle('Ruler');
		expect(r.querySelector('.dve-ruler')).toBeNull();
	});
	afterEach(() => document.body.replaceChildren());

	it('shows the file name and tracks unsaved changes in the title bar', async () => {
		const editor = mount();
		const bar = await inner(editor, 'office-ui-title-bar');
		expect(bar.querySelector('.name')?.textContent).toBe('Report.docx');
		expect(bar.querySelector('.status')?.textContent).toBe('Saved');
		const view = (editor as unknown as { view: import('prosemirror-view').EditorView }).view;
		view.dispatch(view.state.tr.insertText('!', 13));
		await root(editor).querySelector<Updating>('office-ui-title-bar')!.updateComplete;
		expect(bar.querySelector('.status')?.textContent).toBe('Unsaved changes');
	});

	it('runs a ribbon command found through "Tell me what you want to do"', async () => {
		const editor = mount();
		const bar = await inner(editor, 'office-ui-title-bar');
		const search = bar.querySelector<HTMLElement & { value: string }>('office-ui-search')!;
		search.value = 'bulleted';
		search.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
		await root(editor).querySelector<Updating>('office-ui-title-bar')!.updateComplete;
		const options = [...bar.querySelectorAll('[role=option]')];
		expect(options.map((option) => option.textContent?.trim())).toEqual(['Bulleted list']);
		search.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, composed: true }),
		);
		expect(editor.documentModel!.blocks[0]).toMatchObject({ numbering: { level: 0 } });
	});

	it('lets hosts take over file commands by cancelling the file-command event', async () => {
		const editor = mount();
		const seen: string[] = [];
		editor.addEventListener('file-command', (event) => {
			seen.push((event as CustomEvent<{ command: string }>).detail.command);
			event.preventDefault();
		});
		const click = vi.spyOn(HTMLInputElement.prototype, 'click');
		root(editor).querySelector<HTMLButtonElement>('.dve-file-tab')!.click();
		(await inner(editor, 'office-ui-backstage'))
			.querySelector<HTMLButtonElement>('[data-backstage-item="open"]')!
			.click();
		root(editor)
			.querySelector<HTMLButtonElement>('[data-backstage-page="open"] .dve-backstage-primary')!
			.click();
		expect(seen).toEqual(['open']);
		expect(click).not.toHaveBeenCalled();
		click.mockRestore();
	});

	it('lists compatibility notes in File > Info and from the status bar', async () => {
		const editor = mount();
		const notes = (await inner(editor, 'office-ui-status-bar')).querySelector<HTMLButtonElement>(
			'[data-id="notes"]',
		)!;
		expect(notes.hidden).toBe(false);
		expect(notes.textContent?.trim()).toBe('1 compatibility note');
		notes.click();
		const backstage = root(editor).querySelector<HTMLElement>('.dve-backstage')!;
		expect(backstage.hidden).toBe(false);
		expect(backstage.textContent).toContain('Pictures are shown as placeholders.');
		expect(backstage.textContent).toContain('Report.docx');
	});

	it('zooms from the status bar and switches to viewing mode from the title bar', async () => {
		const editor = mount();
		const zoom = await inner(editor, 'office-ui-zoom-slider');
		zoom.querySelector<HTMLButtonElement>('[aria-label="Zoom in"]')!.click();
		await root(editor).querySelector<Updating>('office-ui-zoom-slider')!.updateComplete;
		expect(zoom.querySelector('output')?.textContent).toBe('110%');
		const mode = root(editor).querySelector<HTMLSelectElement>('.dve-mode-select')!;
		mode.value = 'viewing';
		mode.dispatchEvent(new Event('change'));
		expect(editor.readOnly).toBe(true);
	});
});
