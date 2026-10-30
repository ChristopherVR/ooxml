import {
	FILE_COMMAND_EVENT,
	documentSaveMeta,
	fileCommandDetail,
	saveToStorage,
	type DocumentStorage,
	type VisioFileCommand,
} from 'ooxml-core/embedding';

/**
 * Announces a File command as the cancelable `file-command` event, the same contract as Word and
 * Excel. Returns false when a host canceled it to handle the command itself.
 */
export function announceFileCommand(host: HTMLElement, command: VisioFileCommand): boolean {
	return host.dispatchEvent(
		new CustomEvent(FILE_COMMAND_EVENT, {
			detail: fileCommandDetail(command),
			bubbles: true,
			composed: true,
			cancelable: true,
		}),
	);
}

/**
 * The extension a drawing is saved with: .vsdm for a macro-enabled drawing or template (its VBA
 * project is kept, never run), .vsdx for everything else (a template saves as a drawing).
 */
export function drawingExtension(fileName: string): 'vsdx' | 'vsdm' {
	return /\.vs[dt]m$/i.test(fileName.trim()) ? 'vsdm' : 'vsdx';
}

/** The name a drawing is saved under in host storage: its own name as a .vsdx (or .vsdm). */
export function storageFileName(fileName: string): string {
	const base = fileName.trim().replace(/\.vs[dt][xm]?$/i, '') || 'Drawing1';
	return `${base}.${drawingExtension(fileName)}`;
}

/** Saves the drawing's bytes to the host's storage; resolves with the name it now has. */
export function saveDrawingToStorage(
	storage: DocumentStorage,
	bytes: Uint8Array,
	fileName: string,
): Promise<string> {
	return saveToStorage(storage, bytes, documentSaveMeta(storageFileName(fileName)));
}

/** File > Open through the host's storage: the chosen drawing as a named File, or undefined. */
export async function openFromStorage(
	storage: DocumentStorage,
	doc: Document,
): Promise<File | undefined> {
	const file = await storage.open?.();
	if (!file) return undefined;
	const FileType = doc.defaultView?.File ?? File;
	return new FileType([new Uint8Array(file.bytes)], file.fileName);
}
