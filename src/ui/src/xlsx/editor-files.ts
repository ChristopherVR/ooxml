/**
 * Workbook I/O for the element: load (xlsx, xlsm, xls, csv through the core's `loadWorkbook`),
 * new, save, download, and the default handlers of the File commands.
 */
import {
	documentSaveMeta,
	saveToStorage,
	type CloudOpenedFile,
	type DocumentSaveReason,
} from 'ooxml-core/embedding';
import { createWorkbook, type Workbook } from 'ooxml-core/xlsx';
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
import { sheetBaseName } from 'ooxml-core/xlsx/ui';
import { confirmDialog } from '../dialog/message-dialogs';
import { printWorkbook } from './print';
import { packagePassword } from './dialogs/package-password';
import type { SaveState } from './title-bar';
import { saveProgressLabel } from '../files/save-progress-label';

/** Chrome callbacks the file actions report to (title bar save state, backstage). */
export interface FileChrome {
	setSaveState(state: SaveState, detail?: string): void;
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
	cloud?: CloudOpenedFile,
): Promise<void> {
	const generation = ++core.loadGeneration;
	core.loading = true;
	core.requestRender();
	try {
		// The file's fonts start downloading while it parses.
		core.webfonts.prefetch(bytes);
		const { loadWorkbook, isOoxmlCryptoError } = await import('ooxml-core/xlsx/load');
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
		// The webfont fallback (`googleFonts`, off by default) waits here, bounded, so the workbook
		// is first shown with its faces; `loading` stays true meanwhile.
		await core.webfonts.prepare(workbook);
		if (generation !== core.loadGeneration) return;
		showWorkbook(core, workbook, fileName || core.fileName || DEFAULT_FILE_NAME, chrome);
		if (cloud) {
			// Set after showing: showing a workbook drops the previous file's storage and lock.
			core.fileStorage = cloud.storage;
			if (!cloud.canEdit) {
				core.setFileReadOnly(true);
				core.shell?.toast(
					core.ctx.t('This workbook is read-only because the file cannot be edited.'),
					'info',
				);
			}
		}
		if (workbook.warnings.length)
			emit(core.host, 'workbook-warning', { warnings: [...workbook.warnings] });
		const notes = formatNotes(workbook, core.ctx.t);
		if (notes[0]) core.shell?.toast(notes[0], 'info');
	} catch (cause) {
		if (generation === core.loadGeneration) throw reportError(core, cause);
	} finally {
		if (generation === core.loadGeneration) {
			core.loading = false;
			core.requestRender();
		}
	}
}

/** Ask, in an in-app dialog, whether unsaved changes may be discarded. */
function confirmDiscard(core: EditorCore, message: string): Promise<boolean> {
	const t = core.ctx.t;
	return confirmDialog({
		title: t('Unsaved changes'),
		message: t(message),
		confirmLabel: t('Discard'),
		cancelLabel: t('Cancel'),
		tone: 'danger',
		host: core.host,
	});
}

/**
 * Opens a file the user picked in a cloud place: the usual discard confirmation, then the normal
 * load path. The workbook keeps `file.storage` for Save and Save As until another one is shown,
 * and is read-only when the file is (the host's own read-only switch still applies on top).
 */
export async function openCloudFile(
	core: EditorCore,
	file: CloudOpenedFile,
	chrome?: FileChrome,
	confirmed?: () => void,
): Promise<boolean> {
	if (
		core.dirty.dirty &&
		!(await confirmDiscard(core, 'Discard unsaved changes and open another workbook?'))
	)
		return false;
	confirmed?.();
	await loadInto(core, file.bytes, file.fileName, chrome, file);
	return core.fileStorage === file.storage;
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
	if (grid?.isEditing() && !grid.commitEdit())
		throw new Error(core.ctx.t('Finish or cancel the current cell edit before saving.'));
	const { saveWorkbook } = await import('ooxml-core/xlsx/load');
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

/** Editors with a save pending; a Save or Save As meanwhile is ignored, not queued. */
const saving = new WeakSet<EditorCore>();

/** True while `core` has a save pending. */
export const isSaving = (core: EditorCore): boolean => saving.has(core);

/**
 * Saves to the host's storage when there is one, else to a download named after the workbook;
 * .xls and .csv sources become .xlsx (Excel's notice). An edit made while the save runs keeps the
 * workbook dirty: the saved bytes do not contain it.
 */
async function saveDownload(
	core: EditorCore,
	chrome?: FileChrome,
	reason: DocumentSaveReason = 'save',
): Promise<void> {
	const source = core.workbook?.format;
	let name = withExtension(core.fileName, saveExtension(core.fileName));
	// A host that removed Save (`save: false`) never gets `storage.save`: Save As then downloads.
	const storage = core.customization.commands.has('file.save')
		? undefined
		: (core.fileStorage ?? core.storage);
	let edited = false;
	const changed = () => (edited = true);
	saving.add(core);
	chrome?.setSaveState('saving');
	try {
		const bytes = await saveBytes(core, 'xlsx');
		core.host.addEventListener('workbook-change', changed);
		if (storage)
			name = await saveToStorage(storage, bytes, {
				...documentSaveMeta(name, reason),
				onProgress: (progress) =>
					chrome?.setSaveState(
						'saving',
						saveProgressLabel(progress, (text) => core.ctx.t(text)),
					),
			});
		else downloadBytes(core.host.ownerDocument, bytes, name);
	} catch (error) {
		chrome?.setSaveState('dirty');
		throw error;
	} finally {
		core.host.removeEventListener('workbook-change', changed);
		saving.delete(core);
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
	if (edited) return chrome?.setSaveState('dirty');
	core.dirty.set(false);
	chrome?.setSaveState(storage ? 'saved' : 'saved-local');
}

export interface FileCommandOptions {
	chrome?: FileChrome;
	openBackstage?(page: 'saveAs'): void;
	/** Share's default: the Share dialog. */
	openShare?(): void;
}

/** The command id (`hiddenActions` / capability id) of each File command. */
const FILE_COMMAND_IDS: Record<FileCommand, string> = {
	new: 'file.new',
	open: 'file.open',
	save: 'file.save',
	saveAs: 'file.save-as',
	export: 'file.export',
	exportCsv: 'file.export-csv',
	print: 'file.print',
	share: 'file.share',
};

/** Runs a File command unless a host cancels the `file-command` event to handle it itself. */
export async function runFileCommand(
	core: EditorCore,
	command: FileCommand,
	fileName: string | undefined,
	options: FileCommandOptions = {},
): Promise<void> {
	// A command a host capability removed (`save: false` and so on) never runs from any entry
	// point: no `file-command`, no `storage` call.
	if (core.customization.commands.has(FILE_COMMAND_IDS[command])) return;
	if ((command === 'save' || command === 'saveAs') && saving.has(core)) return;
	if (!announceFileCommand(core.host, command, fileName)) return;
	const { chrome } = options;
	const t = core.ctx.t;
	try {
		switch (command) {
			case 'new':
				if (
					core.dirty.dirty &&
					!(await confirmDiscard(core, 'Discard unsaved changes and start a new workbook?'))
				)
					return;
				return newInto(core, chrome);
			case 'open': {
				if (
					core.dirty.dirty &&
					!(await confirmDiscard(core, 'Discard unsaved changes and open another workbook?'))
				)
					return;
				const storage = core.storage;
				const file = storage?.open
					? await storage
							.open()
							.then((opened) => opened && { bytes: opened.bytes, name: opened.fileName })
					: await pickFile(core.host.shadowRoot ?? core.host);
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
				return await saveDownload(core, chrome, 'saveAs');
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
			case 'share':
				return options.openShare?.();
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
