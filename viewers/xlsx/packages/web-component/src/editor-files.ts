/**
 * Workbook I/O for the element: load (xlsx, xlsm, xls, csv through the core's `loadWorkbook`),
 * new, save, download, and the default handlers of the File commands.
 */
import { createWorkbook, type Workbook } from '@christophervr/xlsx-core';
import { createTemplateWorkbook, formatNotes, type TemplateId } from './backstage';
import type { EditorCore } from './editor-core';
import { DEFAULT_FILE_NAME } from './editor-attributes';
import { emit, type FileCommand } from './events';
import {
	announceFileCommand,
	blobFor,
	downloadBytes,
	pickFile,
	saveExtension,
	withExtension,
} from './file-commands';
import { sheetBaseName } from './localization';
import { printWorkbook } from './print';
import { packagePassword } from './dialogs/package-password';
import type { SaveState } from './title-bar';

/** Chrome callbacks the file actions report to (title bar save state, backstage). */
export interface FileChrome {
	setSaveState(state: SaveState): void;
	fileNameChanged(): void;
}

const toError = (cause: unknown) => (cause instanceof Error ? cause : new Error(String(cause)));

export function reportError(core: EditorCore, cause: unknown): Error {
	const error = toError(cause);
	emit(core.host, 'workbook-error', { error, message: error.message });
	core.shell?.toast(error.message, 'error');
	return error;
}

/** Shows a new workbook model: fresh session, clean state, `workbook-change`. */
export function showWorkbook(
	core: EditorCore,
	workbook: Workbook,
	fileName: string,
	chrome?: FileChrome,
): void {
	core.setWorkbook(workbook);
	core.fileName = fileName;
	chrome?.fileNameChanged();
	chrome?.setSaveState('saved');
	emit(core.host, 'workbook-change', { workbook });
}

export async function loadInto(
	core: EditorCore,
	bytes: Uint8Array | ArrayBuffer,
	fileName: string | undefined,
	chrome?: FileChrome,
): Promise<void> {
	const generation = ++core.loadGeneration;
	try {
		const { loadWorkbook, isOoxmlCryptoError } = await import('@christophervr/xlsx-core/load');
		let password: string | undefined;
		let workbook: Workbook;
		for (;;) {
			try {
				workbook = await loadWorkbook(bytes, {
					...(fileName ? { fileName } : {}),
					...(password === undefined ? {} : { password }),
				});
				break;
			} catch (error) {
				if (generation !== core.loadGeneration) return;
				if (!isOoxmlCryptoError(error) || error.code === 'data-integrity' || !core.host.isConnected)
					throw error;
				if (error.code === 'incorrect-password')
					core.ctx.toast(core.ctx.t(error.message), 'warning');
				password = await packagePassword(core.ctx);
				if (password === undefined || generation !== core.loadGeneration) return;
			}
		}
		if (generation !== core.loadGeneration) return;
		showWorkbook(core, workbook, fileName || core.fileName || DEFAULT_FILE_NAME, chrome);
		if (workbook.warnings.length)
			emit(core.host, 'workbook-warning', { warnings: [...workbook.warnings] });
		const notes = formatNotes(workbook, core.ctx.t);
		if (notes[0]) core.shell?.toast(notes[0], 'info');
	} catch (cause) {
		if (generation === core.loadGeneration) throw reportError(core, cause);
	}
}

export function newInto(core: EditorCore, chrome?: FileChrome): void {
	core.loadGeneration++;
	const first = `${sheetBaseName(core.ctx.t)}1`;
	showWorkbook(core, createWorkbook({ sheets: [first] }), DEFAULT_FILE_NAME, chrome);
}

export function templateInto(core: EditorCore, id: TemplateId, chrome?: FileChrome): void {
	core.loadGeneration++;
	const workbook = createTemplateWorkbook(id, core.ctx.t);
	showWorkbook(
		core,
		workbook,
		withExtension(
			core.ctx.t(id === 'todo' ? 'To-do list' : id === 'budget' ? 'Monthly budget' : 'Invoice'),
			'xlsx',
		),
		chrome,
	);
}

export async function saveBytes(
	core: EditorCore,
	format: 'xlsx' | 'csv' = 'xlsx',
): Promise<Uint8Array> {
	if (!core.workbook) throw new Error('No workbook is open');
	const grid = core.ctx.grid();
	if (grid?.isEditing()) grid.commitEdit();
	const { saveWorkbook } = await import('@christophervr/xlsx-core/load');
	return saveWorkbook(core.workbook, format, {
		sheetIndex: core.activeSheet,
		...(format === 'xlsx' && core.savePassword !== undefined
			? { password: core.savePassword }
			: {}),
	});
}

export async function saveBlob(core: EditorCore): Promise<Blob> {
	return blobFor(
		await saveBytes(core, 'xlsx'),
		withExtension(core.fileName, saveExtension(core.fileName)),
	);
}

/** Saves to a download named after the workbook; .xls and .csv sources become .xlsx (Excel's notice). */
async function saveDownload(core: EditorCore, chrome?: FileChrome): Promise<void> {
	const source = core.workbook?.format;
	const name = withExtension(core.fileName, saveExtension(core.fileName));
	chrome?.setSaveState('saving');
	try {
		downloadBytes(core.host.ownerDocument, await saveBytes(core, 'xlsx'), name);
	} catch (error) {
		chrome?.setSaveState('dirty');
		throw error;
	}
	if (name !== core.fileName) {
		core.fileName = name;
		chrome?.fileNameChanged();
		if (source === 'xls' || source === 'csv')
			core.shell?.toast(
				core.ctx.t(
					'Saved as an Excel Workbook (.xlsx): {name}. Some features of the original format may differ.',
					{ name },
				),
				'info',
			);
	}
	core.dirty.set(false);
	chrome?.setSaveState('saved-local');
}

export interface FileCommandOptions {
	chrome?: FileChrome;
	openBackstage?(page: 'saveAs'): void;
}

/** Runs a File command unless a host cancels the `file-command` event to handle it itself. */
export async function runFileCommand(
	core: EditorCore,
	command: FileCommand,
	fileName: string | undefined,
	options: FileCommandOptions = {},
): Promise<void> {
	if (!announceFileCommand(core.host, command, fileName)) return;
	const { chrome } = options;
	const t = core.ctx.t;
	try {
		switch (command) {
			case 'new':
				if (
					core.dirty.dirty &&
					typeof confirm === 'function' &&
					!confirm(t('Discard unsaved changes and start a new workbook?'))
				)
					return;
				return newInto(core, chrome);
			case 'open': {
				if (
					core.dirty.dirty &&
					typeof confirm === 'function' &&
					!confirm(t('Discard unsaved changes and open another workbook?'))
				)
					return;
				const file = await pickFile(core.host.shadowRoot ?? core.host);
				if (file) await loadInto(core, file.bytes, file.name, chrome);
				return;
			}
			case 'save':
				return await saveDownload(core, chrome);
			case 'saveAs': {
				const name = (fileName ?? '').trim();
				if (!name) return options.openBackstage?.('saveAs');
				core.fileName = withExtension(name, saveExtension(name));
				chrome?.fileNameChanged();
				return await saveDownload(core, chrome);
			}
			case 'export':
				return downloadBytes(
					core.host.ownerDocument,
					await saveBytes(core, 'xlsx'),
					withExtension(fileName ?? core.fileName, 'xlsx'),
				);
			case 'exportCsv': {
				const name = withExtension(fileName ?? core.fileName, 'csv');
				downloadBytes(core.host.ownerDocument, await saveBytes(core, 'csv'), name);
				if ((core.workbook?.sheets.length ?? 0) > 1)
					core.shell?.toast(t('Only the current sheet was saved as CSV.'), 'info');
				return;
			}
			case 'print':
				if (!core.workbook) return;
				return printWorkbook(
					core.host.ownerDocument,
					core.workbook,
					core.activeSheet,
					core.fileName,
				);
		}
	} catch (cause) {
		reportError(core, cause);
	}
}
