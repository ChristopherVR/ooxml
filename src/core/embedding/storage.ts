/**
 * Host-owned storage: the contract a host application implements to take over where an embedded
 * editor reads, saves and keeps recovery copies of its document.
 *
 * Every product element (`<docx-editor>`, `<xlsx-editor>`, `<visio-viewer>` and the PowerPoint
 * viewer of each binding) accepts a {@link DocumentStorage} as its `storage` property. When it is
 * set, File > Save hands the saved bytes to {@link DocumentStorage.save} instead of downloading
 * them, File > Open asks {@link DocumentStorage.open} instead of showing the browser's file picker,
 * and an editor with AutoRecover (PowerPoint, Visio) keeps its recovery copies through the
 * recovery methods instead of the library's own IndexedDB database. The cancelable
 * `file-command` event still fires first: a host that prevents it handles the command itself and
 * the storage is not called.
 *
 * DOM-free and format-neutral: the editors only pass bytes, a file name and a reason.
 */

/**
 * A command of the File backstage, title bar or a shortcut, announced as the cancelable
 * `file-command` event before the editor runs its default. Each product raises the subset it has
 * (see {@link WordFileCommand}, {@link ExcelFileCommand}, {@link VisioFileCommand}).
 */
export type FileCommand =
	| 'new'
	| 'open'
	| 'save'
	| 'saveAs'
	| 'export'
	| 'exportText'
	| 'exportCsv'
	| 'print'
	| 'share';

/** The File commands `<docx-editor>` announces. */
export type WordFileCommand = Exclude<FileCommand, 'exportCsv'>;
/** The File commands `<xlsx-editor>` announces. */
export type ExcelFileCommand = Exclude<FileCommand, 'exportText' | 'share'>;
/** The File commands `<visio-viewer>` announces. */
export type VisioFileCommand = Extract<FileCommand, 'new' | 'open' | 'save'>;

/** The detail of a `file-command` event. */
export interface FileCommandDetail<C extends string = FileCommand> {
	command: C;
	/** Save As and the exports: the name the user chose, when there is one. */
	fileName?: string;
}

/** The event name every editor uses for {@link FileCommandDetail}. */
export const FILE_COMMAND_EVENT = 'file-command';

/** Builds a {@link FileCommandDetail}, leaving `fileName` out rather than undefined. */
export function fileCommandDetail<C extends string>(
	command: C,
	fileName?: string,
): FileCommandDetail<C> {
	return fileName === undefined ? { command } : { command, fileName };
}

/** Why an editor is saving: File > Save (or the Quick Access button, or Ctrl+S) or Save As. */
export type DocumentSaveReason = 'save' | 'saveAs';

/** What an editor tells the host about the bytes it hands to {@link DocumentStorage.save}. */
export interface DocumentSaveMeta {
	/** The file name the editor would download as, extension included (`Report.docx`). */
	fileName: string;
	/** The MIME type of the bytes. */
	mimeType: string;
	reason: DocumentSaveReason;
}

/** What {@link DocumentStorage.save} may report back. */
export interface DocumentSaveResult {
	/** The name the document now has, when the host chose another (a Save As dialog of its own). */
	fileName?: string;
}

/** A document the host opened for the editor. */
export interface DocumentOpenResult {
	bytes: Uint8Array;
	fileName: string;
}

/** A recovery copy of a document with unsaved changes. */
export interface DocumentRecoveryEntry {
	/** The document id it is kept under (the editors use the file name or the host's file path). */
	id: string;
	/** When it was written, in milliseconds since the epoch. */
	savedAt: number;
	bytes: Uint8Array;
}

/** A recovery copy without its bytes, as {@link DocumentStorage.listRecovery} returns it. */
export interface DocumentRecoveryInfo {
	id: string;
	savedAt: number;
	/** Size of the copy in bytes. */
	size: number;
}

/**
 * Storage a host supplies to an editor. Only {@link save} is required; each optional method the
 * host leaves out keeps the editor's own behaviour for it (the browser file picker for Open, the
 * library's IndexedDB for recovery copies). The recovery methods are used together: a host
 * provides `readRecovery`, `writeRecovery` and `clearRecovery`, or none of them.
 *
 * Every method may reject; the editor reports a failed save as an error and treats a failed
 * recovery read or write as "no copy kept".
 */
export interface DocumentStorage {
	/** Store the saved document. Resolve once it is stored; reject to report the save as failed. */
	save(bytes: Uint8Array, meta: DocumentSaveMeta): Promise<DocumentSaveResult | void>;
	/** File > Open: resolve with the document to open, or undefined when the user canceled. */
	open?(): Promise<DocumentOpenResult | undefined>;
	/** The recovery copy kept under `id`, if any. */
	readRecovery?(id: string): Promise<DocumentRecoveryEntry | undefined>;
	/** Keep `entry` as the recovery copy of its id, replacing an earlier one. */
	writeRecovery?(entry: DocumentRecoveryEntry): Promise<void>;
	/** Remove the recovery copy kept under `id` (the document was saved or the copy discarded). */
	clearRecovery?(id: string): Promise<void>;
	/** Every kept recovery copy, for a version-history or recovery list. Optional even then. */
	listRecovery?(): Promise<DocumentRecoveryInfo[]>;
}

/** The recovery methods of a storage that provides all three of them. */
export type DocumentRecoveryStore = Required<
	Pick<DocumentStorage, 'readRecovery' | 'writeRecovery' | 'clearRecovery'>
> &
	Pick<DocumentStorage, 'listRecovery'>;

/** The storage's recovery store, or undefined when it does not provide one (use the default). */
export function recoveryStoreOf(
	storage: DocumentStorage | null | undefined,
): DocumentRecoveryStore | undefined {
	if (!storage?.readRecovery || !storage.writeRecovery || !storage.clearRecovery) return undefined;
	return storage as DocumentRecoveryStore;
}

/** True when `value` has the shape of a {@link DocumentStorage} (a `save` function). */
export function isDocumentStorage(value: unknown): value is DocumentStorage {
	return (
		typeof value === 'object' &&
		value !== null &&
		typeof (value as { save?: unknown }).save === 'function'
	);
}

const MIME_TYPES: Readonly<Record<string, string>> = {
	docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
	docm: 'application/vnd.ms-word.document.macroEnabled.12',
	dotx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template',
	doc: 'application/msword',
	xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
	xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12',
	xltx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template',
	xls: 'application/vnd.ms-excel',
	csv: 'text/csv',
	pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
	pptm: 'application/vnd.ms-powerpoint.presentation.macroEnabled.12',
	ppsx: 'application/vnd.openxmlformats-officedocument.presentationml.slideshow',
	potx: 'application/vnd.openxmlformats-officedocument.presentationml.template',
	ppt: 'application/vnd.ms-powerpoint',
	vsdx: 'application/vnd.ms-visio.drawing',
	vsdm: 'application/vnd.ms-visio.drawing.macroEnabled.12',
	vsd: 'application/vnd.visio',
	txt: 'text/plain',
	json: 'application/json',
};

/** The MIME type of an Office file name by its extension; `application/octet-stream` otherwise. */
export function documentMimeType(fileName: string): string {
	const extension = /\.([^./\\]+)$/.exec(fileName)?.[1]?.toLowerCase() ?? '';
	return MIME_TYPES[extension] ?? 'application/octet-stream';
}

/** The {@link DocumentSaveMeta} for saving `fileName`, its MIME type taken from the extension. */
export function documentSaveMeta(
	fileName: string,
	reason: DocumentSaveReason = 'save',
	mimeType: string = documentMimeType(fileName),
): DocumentSaveMeta {
	return { fileName, mimeType, reason };
}

/**
 * Hands `bytes` to the storage and returns the file name the document has afterwards: the one the
 * host reported, else `meta.fileName`. Rejects when the storage does.
 */
export async function saveToStorage(
	storage: DocumentStorage,
	bytes: Uint8Array,
	meta: DocumentSaveMeta,
): Promise<string> {
	const result = await storage.save(bytes, meta);
	const name = result?.fileName?.trim();
	return name ? name : meta.fileName;
}
