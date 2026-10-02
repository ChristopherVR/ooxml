/**
 * File commands raised by the title bar, the File backstage and shortcuts. Each is first announced
 * as a cancelable `file-command` event so hosts can take over file I/O; when no host cancels it the
 * editor runs a browser-only default (file picker, download, print).
 */
import { emit, type FileCommand } from './events';

export const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const TYPES: Record<string, string> = {
	xlsx: XLSX_TYPE,
	xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12',
	csv: 'text/csv;charset=utf-8',
	xls: 'application/vnd.ms-excel',
};

/** File types the Open picker accepts. */
export const OPEN_ACCEPT = '.xlsx,.xlsm,.xltx,.xls,.csv,.tsv,.txt';

/** Dispatches `file-command`; false when a host canceled it. */
export function announceFileCommand(
	host: HTMLElement,
	command: FileCommand,
	fileName?: string,
): boolean {
	return emit(host, 'file-command', { command, ...(fileName === undefined ? {} : { fileName }) });
}

export const extensionOf = (fileName: string): string =>
	/\.([^./\\]+)$/.exec(fileName)?.[1]?.toLowerCase() ?? '';

/** Replaces the extension of `fileName` (or appends one). */
export function withExtension(fileName: string, extension: string): string {
	const base = fileName.trim().replace(/\.[^./\\]+$/, '') || 'Book1';
	return `${base}.${extension}`;
}

/** The extension a save writes: macro workbooks keep .xlsm, everything else (.xls, .csv) is .xlsx. */
export function saveExtension(fileName: string): 'xlsx' | 'xlsm' {
	return extensionOf(fileName) === 'xlsm' ? 'xlsm' : 'xlsx';
}

export function blobFor(bytes: Uint8Array, fileName: string): Blob {
	return new Blob([new Uint8Array(bytes)], { type: TYPES[extensionOf(fileName)] ?? XLSX_TYPE });
}

/** Starts a browser download of `bytes` named `fileName`. */
export function downloadBytes(doc: Document, bytes: Uint8Array, fileName: string): void {
	if (typeof URL.createObjectURL !== 'function') return;
	const url = URL.createObjectURL(blobFor(bytes, fileName));
	const anchor = doc.createElement('a');
	anchor.href = url;
	anchor.download = fileName;
	anchor.style.display = 'none';
	doc.body.append(anchor);
	anchor.click();
	anchor.remove();
	setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/** Opens the browser file picker; resolves with the chosen file's bytes and name. */
export function pickFile(
	root: ShadowRoot | HTMLElement,
	accept = OPEN_ACCEPT,
): Promise<{ bytes: Uint8Array; name: string } | undefined> {
	const doc = root.ownerDocument ?? document;
	const input = doc.createElement('input');
	input.type = 'file';
	input.accept = accept;
	input.hidden = true;
	root.append(input);
	return new Promise((resolve, reject) => {
		input.addEventListener('change', () => {
			const file = input.files?.[0];
			input.remove();
			if (!file) return resolve(undefined);
			file
				.arrayBuffer()
				.then((buffer) => resolve({ bytes: new Uint8Array(buffer), name: file.name }), reject);
		});
		input.addEventListener('cancel', () => {
			input.remove();
			resolve(undefined);
		});
		input.click();
	});
}
