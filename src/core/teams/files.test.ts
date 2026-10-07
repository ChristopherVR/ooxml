import { createFileActions } from './files.js';
import { loadXlsx } from '../xlsx/index.js';
import type { Attachment } from './model.js';

function setup() {
	const uploads: (Blob & { name: string })[] = [];
	const messages: { channelId: string; text: string; attachments: Attachment[] }[] = [];
	const available = vi.fn(() => true);
	const canUpload = vi.fn(() => true);
	const upload = vi.fn(async (file: Blob & { name: string }) => {
		uploads.push(file);
		return {
			name: file.name,
			kind: 'other' as const,
			mime: file.type,
			url: `https://files.test/${file.name}`,
		};
	});
	const actions = createFileActions({
		available,
		canUpload,
		upload,
		post: (channelId, text, attachments) => {
			messages.push({ channelId, text, attachments });
			return true;
		},
	});
	return { ...actions, uploads, messages, available, canUpload, upload };
}

describe('channel file actions', () => {
	it('creates valid native workbook bytes and never overwrites an existing name', async () => {
		const actions = setup();
		const first = await actions.createWorkbook('finance', 'Budget.xlsx');
		const second = await actions.createWorkbook('finance', 'Budget');
		expect(first.name).toBe('Budget.xlsx');
		expect(first.kind).toBe('xlsx');
		expect(first.url).not.toBe(second.url);
		const workbook = await loadXlsx(new Uint8Array(await actions.uploads[0]!.arrayBuffer()));
		expect(workbook.sheets.map((sheet) => sheet.name)).toEqual(['Sheet1']);
		expect(actions.messages[0]).toMatchObject({
			channelId: 'finance',
			text: 'Created Budget.xlsx',
		});
	});
	it('shares duplicate upload names at separate URLs and keeps their MIME types', async () => {
		const actions = setup();
		const file = Object.assign(new Blob(['# Notes'], { type: 'text/markdown' }), {
			name: 'Notes.md',
		});
		const attachments = await actions.uploadFiles('project', [file, file]);
		expect(attachments.map((attachment) => attachment.name)).toEqual(['Notes.md', 'Notes.md']);
		expect(attachments[0]!.url).not.toBe(attachments[1]!.url);
		expect(attachments[0]!.mime).toBe('text/markdown');
		expect(actions.messages[0]!.channelId).toBe('project');
	});
	it('rejects missing storage, unsafe upload results and invalid batches without sharing broken links', async () => {
		const actions = setup();
		actions.canUpload.mockReturnValue(false);
		await expect(actions.createWorkbook('finance', 'Budget')).rejects.toThrow(
			'Configure file storage',
		);
		expect(actions.upload).not.toHaveBeenCalled();
		actions.canUpload.mockReturnValue(true);
		await expect(actions.uploadFiles('finance', [])).rejects.toThrow('Choose between');
		await expect(
			actions.uploadFiles('finance', [Object.assign(new Blob(), { name: '' })]),
		).rejects.toThrow('cannot be saved');
		actions.upload.mockResolvedValue({
			name: 'x',
			kind: 'other',
			mime: '',
			url: 'https://user:secret@host/file',
		});
		await expect(
			actions.saveFileCopy('finance', Object.assign(new Blob(['x']), { name: 'Budget.xlsx' })),
		).rejects.toThrow('could not be uploaded');
		expect(actions.messages).toEqual([]);
	});
	it('does not publish to an archived channel or post partial batches after an upload failure', async () => {
		const actions = setup();
		const file = Object.assign(new Blob(['x']), { name: 'Budget.xlsx' });
		actions.upload.mockImplementationOnce(async (file) => {
			actions.available.mockReturnValue(false);
			return { name: file.name, kind: 'other', mime: '', url: 'https://files.test/uploaded.xlsx' };
		});
		await expect(actions.uploadFiles('finance', [file])).rejects.toThrow('not shared');
		expect(actions.messages).toEqual([]);
		actions.available.mockReturnValue(true);
		actions.upload.mockResolvedValueOnce({
			name: 'file.xlsx',
			kind: 'other',
			mime: '',
			url: 'https://files.test/first.xlsx',
		});
		actions.upload.mockRejectedValueOnce(new Error('Storage offline'));
		await expect(actions.uploadFiles('finance', [file, file])).rejects.toThrow('Storage offline');
		expect(actions.messages).toEqual([]);
	});
});
