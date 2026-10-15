import type JSZip from 'jszip';
import { openZip } from '../../../opc/zip-limits';
import { DEFAULT_MAX_UNCOMPRESSED_BYTES, MAX_ZIP_ENTRY_COUNT, ZipBombError } from '../core/types';

const createError = (message: string, details: { limit: number }): Error =>
	new ZipBombError(message, details);

/** Opens a whole presentation package with the PowerPoint size limits (500 MiB inflated). */
export function openPptxZip(bytes: Uint8Array | ArrayBuffer): Promise<JSZip> {
	return openZip(bytes, {
		label: 'PPTX',
		maxInputBytes: DEFAULT_MAX_UNCOMPRESSED_BYTES,
		maxEntries: MAX_ZIP_ENTRY_COUNT,
		maxEntryBytes: DEFAULT_MAX_UNCOMPRESSED_BYTES,
		maxTotalBytes: DEFAULT_MAX_UNCOMPRESSED_BYTES,
		createError,
	});
}

/** Opens a package nested in a presentation (embedded workbook, document, OLE payload). */
export function openNestedZip(bytes: Uint8Array | ArrayBuffer): Promise<JSZip> {
	return openZip(bytes, { label: 'Embedded', createError });
}
