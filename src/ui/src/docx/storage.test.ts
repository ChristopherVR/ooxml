// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createDocument, saveDocx } from 'ooxml-core/docx';
import type { DocumentSaveMeta, DocumentStorage } from 'ooxml-core/embedding';
import type { EditorChrome } from './editor-chrome';
import { DocxEditorElement, registerDocxEditor } from './index';

registerDocxEditor();

function mount(): DocxEditorElement {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	editor.documentModel = createDocument();
	editor.fileName = 'Report.docx';
	document.body.append(editor);
	return editor;
}
const chrome = (editor: DocxEditorElement) =>
	(editor as unknown as { core: { shell: { chrome: EditorChrome } } }).core.shell.chrome;

function recording(result?: { fileName: string }) {
	const saves: Array<[Uint8Array, DocumentSaveMeta]> = [];
	const storage: DocumentStorage = {
		save: async (bytes, meta) => {
			saves.push([bytes, meta]);
			return result;
		},
	};
	return { storage, saves };
}

describe('<docx-editor> host-owned storage', () => {
	afterEach(() => {
		document.body.replaceChildren();
		vi.restoreAllMocks();
	});

	it('hands Save to storage.save instead of downloading', async () => {
		const editor = mount();
		const { storage, saves } = recording();
		editor.storage = storage;
		expect(editor.storage).toBe(storage);
		const download = vi.spyOn(HTMLAnchorElement.prototype, 'click');
		await chrome(editor).run('save');
		expect(saves).toHaveLength(1);
		const [bytes, meta] = saves[0]!;
		expect(bytes.byteLength).toBeGreaterThan(0);
		expect(meta).toEqual({
			fileName: 'Report.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			reason: 'save',
		});
		expect(download).not.toHaveBeenCalled();
		expect(editor.dirty).toBe(false);
	});

	it('reports Save As with the chosen name, and takes the name the host reports back', async () => {
		const editor = mount();
		const { storage, saves } = recording({ fileName: 'Final.docx' });
		editor.storage = storage;
		await chrome(editor).run('saveAs', 'Draft');
		expect(saves[0]![1]).toMatchObject({ fileName: 'Draft.docx', reason: 'saveAs' });
		expect(editor.fileName).toBe('Final.docx');
	});

	it('does not call storage when the host cancels file-command', async () => {
		const editor = mount();
		const { storage, saves } = recording();
		editor.storage = storage;
		editor.addEventListener('file-command', (event) => event.preventDefault());
		await chrome(editor).run('save');
		await chrome(editor).run('saveAs', 'Other');
		expect(saves).toHaveLength(0);
	});

	it('reports a failed storage save as an error and stays dirty', async () => {
		const editor = mount();
		editor.storage = { save: () => Promise.reject(new Error('offline')) };
		const errors: string[] = [];
		editor.addEventListener('document-error', (event) =>
			errors.push(String((event as CustomEvent).detail?.message ?? (event as CustomEvent).detail)),
		);
		await chrome(editor).run('save');
		expect(errors.join()).toContain('offline');
		expect(editor.dirty).toBe(true);
	});

	it('opens through storage.open instead of the file picker', async () => {
		const editor = mount();
		const bytes = await saveDocx(createDocument());
		const click = vi.spyOn(HTMLInputElement.prototype, 'click');
		editor.storage = {
			save: async () => undefined,
			open: async () => ({ bytes, fileName: 'Opened.docx' }),
		};
		await chrome(editor).run('open');
		expect(click).not.toHaveBeenCalled();
		expect(editor.fileName).toBe('Opened.docx');
	});

	it('keeps the file picker when the storage cannot open', async () => {
		const editor = mount();
		const click = vi.spyOn(HTMLInputElement.prototype, 'click').mockImplementation(() => {});
		editor.storage = { save: async () => undefined };
		await chrome(editor).run('open');
		expect(click).toHaveBeenCalled();
	});
});
