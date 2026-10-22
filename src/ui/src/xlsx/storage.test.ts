// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createWorkbook } from 'ooxml-core/xlsx';
import { loadWorkbook, saveWorkbook } from 'ooxml-core/xlsx/load';
import type { DocumentSaveMeta, DocumentStorage } from 'ooxml-core/embedding';
import { runFileCommand } from './editor-files';
import * as fileCommands from './file-commands';
import { shellFixture } from './test-support/shell';
import { XlsxEditorElement, defineXlsxEditor } from './index';

afterEach(() => {
	vi.restoreAllMocks();
	document.body.innerHTML = '';
});

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
const chrome = () => ({ setSaveState: vi.fn(), fileNameChanged: vi.fn() });

describe('<xlsx-editor> host-owned storage', () => {
	it('is a property of the element', () => {
		defineXlsxEditor();
		const editor = document.createElement('xlsx-editor') as XlsxEditorElement;
		const { storage } = recording();
		editor.storage = storage;
		expect(editor.storage).toBe(storage);
		editor.storage = null;
		expect(editor.storage).toBeUndefined();
	});

	it('hands Save to storage.save instead of downloading, and marks the workbook clean', async () => {
		const { core } = shellFixture();
		const { storage, saves } = recording();
		core.storage = storage;
		core.dirty.set(true);
		const download = vi.spyOn(fileCommands, 'downloadBytes');
		const ui = chrome();
		await runFileCommand(core, 'save', undefined, { chrome: ui });
		expect(download).not.toHaveBeenCalled();
		expect(saves).toHaveLength(1);
		const [bytes, meta] = saves[0]!;
		await expect(loadWorkbook(bytes)).resolves.toHaveProperty('sheets');
		expect(meta).toEqual({
			fileName: 'Book1.xlsx',
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			reason: 'save',
		});
		expect(core.dirty.dirty).toBe(false);
		expect(ui.setSaveState.mock.calls.map(([state]) => state)).toEqual(['saving', 'saved']);
	});

	it('reports Save As and takes the name the host reports back', async () => {
		const { core } = shellFixture();
		const { storage, saves } = recording({ fileName: 'Budget final.xlsx' });
		core.storage = storage;
		await runFileCommand(core, 'saveAs', 'Budget', { chrome: chrome() });
		expect(saves[0]![1]).toMatchObject({ fileName: 'Budget.xlsx', reason: 'saveAs' });
		expect(core.fileName).toBe('Budget final.xlsx');
	});

	it('does not call storage when the host cancels file-command', async () => {
		const { core, host } = shellFixture();
		const { storage, saves } = recording();
		core.storage = storage;
		host.addEventListener('file-command', (event) => event.preventDefault());
		await runFileCommand(core, 'save', undefined);
		expect(saves).toHaveLength(0);
	});

	it('reports a failed storage save and stays dirty', async () => {
		const { core, events } = shellFixture();
		core.storage = { save: () => Promise.reject(new Error('offline')) };
		core.dirty.set(true);
		await runFileCommand(core, 'save', undefined, { chrome: chrome() });
		expect(events.filter((event) => event.type === 'workbook-error')).toHaveLength(1);
		expect(core.dirty.dirty).toBe(true);
	});

	it('ignores a second Save while one is pending and stays dirty after a concurrent edit', async () => {
		const { core } = shellFixture();
		let resolve!: () => void;
		const save = vi.fn(() => new Promise<void>((done) => (resolve = done)));
		core.storage = { save };
		core.dirty.set(true);
		const ui = chrome();
		const first = runFileCommand(core, 'save', undefined, { chrome: ui });
		await vi.waitFor(() => expect(save).toHaveBeenCalledTimes(1));
		await runFileCommand(core, 'save', undefined, { chrome: ui });
		expect(save).toHaveBeenCalledTimes(1);
		core.ctx.session()!.setCellInput(0, 0, 0, '5');
		resolve();
		await first;
		expect(core.dirty.dirty).toBe(true);
		expect(ui.setSaveState.mock.calls.at(-1)?.[0]).toBe('dirty');
		const again = runFileCommand(core, 'save', undefined, { chrome: ui });
		await vi.waitFor(() => expect(save).toHaveBeenCalledTimes(2));
		resolve();
		await again;
		expect(core.dirty.dirty).toBe(false);
	});

	it('opens through storage.open instead of the file picker', async () => {
		const { core } = shellFixture();
		const bytes = await saveWorkbook(createWorkbook({ sheets: ['Data'] }), 'xlsx');
		const pick = vi.spyOn(fileCommands, 'pickFile');
		core.storage = {
			save: async () => undefined,
			open: async () => ({ bytes, fileName: 'Data.xlsx' }),
		};
		await runFileCommand(core, 'open', undefined);
		expect(pick).not.toHaveBeenCalled();
		expect(core.fileName).toBe('Data.xlsx');
		expect(core.workbook?.sheets[0]?.name).toBe('Data');
	});

	it('does nothing when storage.open is canceled', async () => {
		const { core } = shellFixture();
		const before = core.workbook;
		core.storage = { save: async () => undefined, open: async () => undefined };
		await runFileCommand(core, 'open', undefined);
		expect(core.workbook).toBe(before);
	});
});
