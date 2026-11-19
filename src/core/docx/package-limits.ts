import type { ZipLimits } from '../opc/zip-limits';

/**
 * Size limits a DOCX package is read and saved under; each field left out keeps its default. The
 * defaults guard against zip bombs and runaway memory use; a host that expects very large
 * documents raises them explicitly.
 */
export interface DocxLoadLimits {
	/** Largest accepted package, compressed (default 50 MiB). */
	maxInputBytes?: number;
	/** Largest number of parts (zip entries) (default 5,000). */
	maxParts?: number;
	/** Largest inflated size of one part (default 200 MiB). */
	maxPartBytes?: number;
	/** Largest total inflated size of all parts (default 200 MiB). */
	maxTotalBytes?: number;
	/** Largest `word/document.xml`, inflated (default 25 MiB). */
	maxDocumentPartBytes?: number;
}

/** The limits a DOCX load applies unless the caller raises (or lowers) them. */
export const DEFAULT_DOCX_LOAD_LIMITS: Readonly<Required<DocxLoadLimits>> = {
	maxInputBytes: 50 * 1024 * 1024,
	maxParts: 5000,
	maxPartBytes: 200 * 1024 * 1024,
	maxTotalBytes: 200 * 1024 * 1024,
	maxDocumentPartBytes: 25 * 1024 * 1024,
};

/** Resolved limits: the zip guard's and the main document part's. */
export interface ResolvedDocxLimits {
	zip: ZipLimits;
	maxDocumentPartBytes: number;
}

/**
 * The defaults overridden by `limits`. Every given value must be a positive finite number, so a
 * computed value (`NaN`, `Infinity`) cannot switch a guard off by accident.
 */
export function resolveDocxLimits(limits: DocxLoadLimits = {}): ResolvedDocxLimits {
	const resolved = { ...DEFAULT_DOCX_LOAD_LIMITS };
	for (const name of Object.keys(resolved) as (keyof DocxLoadLimits)[]) {
		const value = limits[name];
		if (value === undefined) continue;
		if (!Number.isFinite(value) || value <= 0)
			throw new RangeError(`limits.${name} must be a positive finite number`);
		resolved[name] = value;
	}
	return {
		zip: {
			maxInputBytes: resolved.maxInputBytes,
			maxEntries: resolved.maxParts,
			maxEntryBytes: resolved.maxPartBytes,
			maxTotalBytes: resolved.maxTotalBytes,
			label: 'DOCX',
		},
		maxDocumentPartBytes: resolved.maxDocumentPartBytes,
	};
}
