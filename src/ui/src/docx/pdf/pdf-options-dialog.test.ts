// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createDocument } from 'ooxml-core/docx';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { DocxEditorElement, registerDocxEditor } from '../index';
import { ribbonFile } from '../test-support';
import { createPdfOptionsDialog } from './pdf-options-dialog';
import type { DocxPdfResult } from './export-pdf';

registerDocxEditor();
afterEach(() => document.body.replaceChildren());

const input = (root: ParentNode, label: string) =>
	root.querySelector<HTMLInputElement>(`input[aria-label="${label}"]`)!;
const button = (root: ParentNode, label: string) =>
	[...root.querySelectorAll('button')].find((b) => b.textContent === label)!;

function setup() {
	const dialog = createPdfOptionsDialog();
	document.body.append(dialog.element);
	return dialog;
}

describe('PDF Options dialog', () => {
	it('opens on the defaults and exports with the chosen options', async () => {
		const dialog = setup();
		const result = dialog.choose({ pages: { current: 2, total: 4 }, hasSelection: true });
		const root = dialog.element;
		expect(dialog.isOpen).toBe(true);
		expect(input(root, 'All').checked).toBe(true);
		expect(input(root, 'Document').checked).toBe(true);
		expect(input(root, 'Create bookmarks using').checked).toBe(true);
		expect(input(root, 'Headings').checked).toBe(true);
		expect(input(root, 'Document properties').checked).toBe(true);
		expect(input(root, 'Document structure tags for accessibility').checked).toBe(true);
		expect(input(root, 'PDF/A compliant').checked).toBe(false);
		expect(input(root, 'From').value).toBe('2');
		const balloons = root.querySelector<HTMLSelectElement>('select[aria-label="Balloons"]')!;
		expect(balloons.disabled).toBe(true);
		input(root, 'Current page').click();
		input(root, 'Document showing markup').click();
		expect(balloons.disabled).toBe(false);
		balloons.value = 'comments';
		input(root, 'Word bookmarks').click();
		input(root, 'Document properties').click();
		input(root, 'Document structure tags for accessibility').click();
		input(root, 'PDF/A compliant').click();
		button(root, 'Export').click();
		expect(await result).toEqual({
			pages: 'current',
			markup: 'all',
			markupBalloons: 'comments',
			bookmarks: 'wordBookmarks',
			documentProperties: false,
			taggedPdf: false,
			pdfA: true,
		});
		expect(dialog.isOpen).toBe(false);
	});

	it('cancels with Cancel or Escape and offers only the ranges the editor knows', async () => {
		const dialog = setup();
		const cancelled = dialog.choose({ hasSelection: false });
		expect(input(dialog.element, 'Current page').disabled).toBe(true);
		expect(input(dialog.element, 'Selection').disabled).toBe(true);
		button(dialog.element, 'Cancel').click();
		expect(await cancelled).toBeUndefined();
		const escaped = dialog.choose({ hasSelection: true });
		expect(input(dialog.element, 'Selection').disabled).toBe(false);
		input(dialog.element, 'All').dispatchEvent(
			new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }),
		);
		expect(await escaped).toBeUndefined();
		expect(dialog.isOpen).toBe(false);
	});

	it('turns the bookmark sources off with the bookmarks and keeps a bad page range open', async () => {
		const dialog = setup();
		const result = dialog.choose({ pages: { current: 1, total: 3 }, hasSelection: false });
		const root = dialog.element;
		input(root, 'Create bookmarks using').click();
		expect(input(root, 'Headings').disabled).toBe(true);
		const from = input(root, 'From');
		const to = input(root, 'To');
		from.value = '3';
		from.dispatchEvent(new Event('input', { bubbles: true }));
		expect(input(root, 'Pages').checked).toBe(true);
		to.value = '2';
		button(root, 'Export').click();
		expect(dialog.isOpen).toBe(true);
		expect(to.getAttribute('aria-invalid')).toBe('true');
		to.value = '9';
		button(root, 'Export').click();
		expect(await result).toMatchObject({ pages: { from: 3, to: 3 }, bookmarks: 'none' });
		// The choices are kept for the next export; the range starts at All again.
		void dialog.choose({ hasSelection: false });
		expect(input(root, 'All').checked).toBe(true);
		expect(input(root, 'Create bookmarks using').checked).toBe(false);
		button(root, 'Cancel').click();
	});

	it('translates its labels', () => {
		const dialog = setup();
		dialog.setLocale('de');
		void dialog.choose({ hasSelection: false });
		expect(dialog.element.textContent).toContain('Seitenbereich');
		expect(button(dialog.element, 'Exportieren')).toBeDefined();
		button(dialog.element, 'Abbrechen').click();
	});
});

describe('File > Export > Create a PDF', () => {
	function mount(): DocxEditorElement {
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		const model = createDocument();
		model.blocks[0] = { type: 'paragraph', id: 'p1', runs: [{ text: 'Hello PDF' }] };
		editor.documentModel = model;
		editor.fileName = 'Report.docx';
		document.body.append(editor);
		return editor;
	}
	type Updating = HTMLElement & { updateComplete: Promise<unknown> };
	async function createPdf(editor: DocxEditorElement) {
		const root = editor.shadowRoot!;
		(await ribbonFile(root)).click();
		const backstage = root.querySelector<Updating>('office-ui-backstage')!;
		await backstage.updateComplete;
		backstage
			.shadowRoot!.querySelector<HTMLButtonElement>('[data-backstage-item="export"]')!
			.click();
		button(root.querySelector('[data-backstage-page="export"]')!, 'Create PDF').click();
		return root.querySelector<HTMLElement>('.dve-pdf-options-dialog')!;
	}

	it('asks for the options, then exports with them; Cancel exports nothing', async () => {
		const editor = mount();
		const bytes = new Uint8Array([37, 80, 68, 70]);
		const exportPdf = vi
			.spyOn(editor, 'exportPdf')
			.mockResolvedValue({ bytes } as unknown as DocxPdfResult);
		const original = { create: URL.createObjectURL, revoke: URL.revokeObjectURL };
		const createObjectURL = vi.fn(() => 'blob:pdf');
		URL.createObjectURL = createObjectURL;
		URL.revokeObjectURL = vi.fn();
		const click = vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {});
		try {
			const view = (editor as unknown as { view: EditorView }).view;
			view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 6)));
			let dialog = await createPdf(editor);
			expect((dialog as HTMLElement & { open: boolean }).open).toBe(true);
			button(dialog, 'Cancel').click();
			await Promise.resolve();
			expect(exportPdf).not.toHaveBeenCalled();
			dialog = await createPdf(editor);
			input(dialog, 'Selection').click();
			input(dialog, 'PDF/A compliant').click();
			vi.useFakeTimers({ toFake: ['setTimeout'] });
			button(dialog, 'Export').click();
			await vi.waitFor(() => expect(click).toHaveBeenCalled());
			// The download revokes its URL a second later.
			vi.advanceTimersByTime(1000);
			expect(exportPdf).toHaveBeenCalledWith(
				expect.objectContaining({ pages: 'selection', pdfA: true, markup: 'none' }),
			);
			expect(createObjectURL).toHaveBeenCalled();
		} finally {
			vi.useRealTimers();
			click.mockRestore();
			exportPdf.mockRestore();
			URL.createObjectURL = original.create;
			URL.revokeObjectURL = original.revoke;
		}
	});
});
