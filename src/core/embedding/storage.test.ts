import { describe, expect, it } from 'vitest';
import {
	documentMimeType,
	documentSaveMeta,
	fileCommandDetail,
	isDocumentStorage,
	recoveryStoreOf,
	saveToStorage,
	type DocumentStorage,
} from './storage';

const noop = async () => undefined;

describe('fileCommandDetail', () => {
	it('leaves fileName out when there is none', () => {
		expect(fileCommandDetail('save')).toStrictEqual({ command: 'save' });
		expect(fileCommandDetail('saveAs', 'A.docx')).toStrictEqual({
			command: 'saveAs',
			fileName: 'A.docx',
		});
	});
});

describe('documentMimeType', () => {
	it('maps every product extension, case-insensitively', () => {
		expect(documentMimeType('Report.DOCX')).toBe(
			'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		);
		expect(documentMimeType('Book1.xlsm')).toBe('application/vnd.ms-excel.sheet.macroEnabled.12');
		expect(documentMimeType('deck.ppt')).toBe('application/vnd.ms-powerpoint');
		expect(documentMimeType('Drawing1.vsdx')).toBe('application/vnd.ms-visio.drawing');
	});
	it('falls back to octet-stream for unknown or missing extensions', () => {
		expect(documentMimeType('notes')).toBe('application/octet-stream');
		expect(documentMimeType('a.zip')).toBe('application/octet-stream');
	});
});

describe('documentSaveMeta', () => {
	it('defaults the reason to save and takes the MIME type from the name', () => {
		expect(documentSaveMeta('Book1.xlsx')).toStrictEqual({
			fileName: 'Book1.xlsx',
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			reason: 'save',
		});
		expect(documentSaveMeta('x.bin', 'saveAs', 'application/x-test').mimeType).toBe(
			'application/x-test',
		);
	});
});

describe('recoveryStoreOf', () => {
	it('needs all three recovery methods', () => {
		expect(recoveryStoreOf(undefined)).toBeUndefined();
		expect(recoveryStoreOf({ save: noop })).toBeUndefined();
		expect(
			recoveryStoreOf({ save: noop, readRecovery: noop, writeRecovery: noop }),
		).toBeUndefined();
		const full: DocumentStorage = {
			save: noop,
			readRecovery: noop,
			writeRecovery: noop,
			clearRecovery: noop,
		};
		expect(recoveryStoreOf(full)).toBe(full);
	});
});

describe('isDocumentStorage', () => {
	it('accepts anything with a save function', () => {
		expect(isDocumentStorage({ save: noop })).toBe(true);
		expect(isDocumentStorage({ save: 1 })).toBe(false);
		expect(isDocumentStorage(null)).toBe(false);
	});
});

describe('saveToStorage', () => {
	it('hands the bytes and meta over and keeps the name when the host reports none', async () => {
		const calls: unknown[] = [];
		const storage: DocumentStorage = {
			save: async (bytes, meta) => void calls.push([bytes, meta]),
		};
		const bytes = new Uint8Array([1, 2]);
		const meta = documentSaveMeta('A.docx');
		expect(await saveToStorage(storage, bytes, meta)).toBe('A.docx');
		expect(calls).toStrictEqual([[bytes, meta]]);
	});
	it('returns the name the host chose, ignoring a blank one', async () => {
		const renamed: DocumentStorage = { save: async () => ({ fileName: ' B.docx ' }) };
		expect(await saveToStorage(renamed, new Uint8Array(), documentSaveMeta('A.docx'))).toBe(
			'B.docx',
		);
		const blank: DocumentStorage = { save: async () => ({ fileName: '  ' }) };
		expect(await saveToStorage(blank, new Uint8Array(), documentSaveMeta('A.docx'))).toBe('A.docx');
	});
	it('rejects when the storage does', async () => {
		const failing: DocumentStorage = {
			save: () => Promise.reject(new Error('offline')),
		};
		await expect(saveToStorage(failing, new Uint8Array(), documentSaveMeta('A'))).rejects.toThrow(
			'offline',
		);
	});
});
