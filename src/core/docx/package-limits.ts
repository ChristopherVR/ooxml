import type { ZipLimits } from '../opc/zip-limits';

/** Limits applied whenever a DOCX package is opened (read or save). */
export const DOCX_ZIP_LIMITS: ZipLimits = {
	maxInputBytes: 50 * 1024 * 1024,
	maxEntries: 5000,
	maxEntryBytes: 200 * 1024 * 1024,
	maxTotalBytes: 200 * 1024 * 1024,
	label: 'DOCX',
};

/** Largest accepted `word/document.xml`. */
export const MAX_DOCUMENT_PART_BYTES = 25 * 1024 * 1024;
