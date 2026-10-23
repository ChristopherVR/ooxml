// Size limits applied while reading a workbook package. The defaults guard against zip bombs and
// runaway memory use; a caller that expects very large workbooks raises them explicitly.
import { ZipLimitError, type ZipLimitKind, type ZipLimits } from '../../opc/index';
import { checkXmlLimits, parseXml, XmlLimitError, type XmlDocument } from '../../xml/index';

/** Limits on a workbook package; each field left out keeps its default. */
export interface XlsxLoadLimits {
	/** Largest accepted package, compressed (default 50 MiB). */
	maxInputBytes?: number;
	/** Largest number of parts (zip entries) (default 10,000). */
	maxParts?: number;
	/** Largest inflated size of one part (default 200 MiB). */
	maxPartBytes?: number;
	/** Largest total inflated size of all parts (default 300 MiB). */
	maxTotalBytes?: number;
	/**
	 * Largest number of XML elements in the shared strings, or in a worksheet whose cell data has to
	 * be parsed with the rest of the part (default 5,000,000). Other parts keep the XML default.
	 */
	maxXmlElements?: number;
}

/** The limits a workbook load applies unless the caller raises (or lowers) them. */
export const DEFAULT_XLSX_LOAD_LIMITS: Readonly<Required<XlsxLoadLimits>> = {
	maxInputBytes: 50 * 1024 * 1024,
	maxParts: 10_000,
	maxPartBytes: 200 * 1024 * 1024,
	maxTotalBytes: 300 * 1024 * 1024,
	maxXmlElements: 5_000_000,
};

/**
 * A workbook (or one of its parts) is larger than a load limit allows. Each entry point bundles
 * its own copy of this class, so test with {@link isWorkbookTooLargeError} (or `code`) rather
 * than `instanceof` when the error may come from another entry point.
 */
export class WorkbookTooLargeError extends Error {
	readonly code = 'workbook-too-large';
	/** The limit that was exceeded, as named in {@link XlsxLoadLimits}. */
	readonly limit: keyof XlsxLoadLimits;
	/** The value of that limit when the load failed. */
	readonly maximum: number;
	constructor(
		message: string,
		limit: keyof XlsxLoadLimits,
		maximum: number,
		options?: { cause?: unknown },
	) {
		super(message, options);
		this.name = 'WorkbookTooLargeError';
		this.limit = limit;
		this.maximum = maximum;
	}
}

/** Whether `error` is a {@link WorkbookTooLargeError} (by `code`, across bundles). */
export function isWorkbookTooLargeError(error: unknown): error is WorkbookTooLargeError {
	return error instanceof Error && (error as { code?: unknown }).code === 'workbook-too-large';
}

/**
 * The defaults overridden by `limits`. Every given value must be a positive finite number, so a
 * computed value (`NaN`, `Infinity`) cannot switch a guard off by accident.
 */
export function resolveLoadLimits(limits: XlsxLoadLimits = {}): Required<XlsxLoadLimits> {
	const resolved = { ...DEFAULT_XLSX_LOAD_LIMITS };
	for (const name of Object.keys(resolved) as (keyof XlsxLoadLimits)[]) {
		const value = limits[name];
		if (value === undefined) continue;
		if (!Number.isFinite(value) || value <= 0)
			throw new RangeError(`limits.${name} must be a positive finite number`);
		resolved[name] = value;
	}
	return resolved;
}

/** The workbook limit behind each size check of the zip guard. */
const ZIP_LIMIT_NAMES: Partial<Record<ZipLimitKind, keyof XlsxLoadLimits>> = {
	maxInputBytes: 'maxInputBytes',
	maxEntries: 'maxParts',
	maxEntryBytes: 'maxPartBytes',
	maxTotalBytes: 'maxTotalBytes',
};

/**
 * The zip guard settings for `limits`. Size violations become {@link WorkbookTooLargeError};
 * forged sizes and unsafe part names stay `ZipLimitError`, since no limit lets them through.
 */
export function zipLimitsFor(limits: Required<XlsxLoadLimits>): ZipLimits {
	return {
		maxInputBytes: limits.maxInputBytes,
		maxEntries: limits.maxParts,
		maxEntryBytes: limits.maxPartBytes,
		maxTotalBytes: limits.maxTotalBytes,
		label: 'XLSX',
		createError: (message, { limit, kind }) => {
			const name = ZIP_LIMIT_NAMES[kind];
			return name
				? new WorkbookTooLargeError(message, name, limit)
				: new ZipLimitError(message, { limit, kind });
		},
	};
}

/** The element-count limit of the XML parser, reported as the `maxXmlElements` load limit. */
function workbookLimit(error: unknown): unknown {
	return error instanceof XmlLimitError && error.limit === 'maxNodes'
		? new WorkbookTooLargeError(error.message, 'maxXmlElements', error.maximum, { cause: error })
		: error;
}

/** Parses a part whose size follows the cell data under the `maxXmlElements` limit. */
export function parseDataXml(
	xml: string,
	label: string,
	maxXmlElements = DEFAULT_XLSX_LOAD_LIMITS.maxXmlElements,
): XmlDocument {
	try {
		return parseXml(xml, { label, limits: { maxNodes: maxXmlElements } });
	} catch (error) {
		throw workbookLimit(error);
	}
}

/**
 * Checks such a part against the element-count and nesting limits {@link parseDataXml} applies,
 * throwing the same errors, without building a DOM (for a part read by the light reader).
 */
export function checkDataXml(
	xml: string,
	label: string,
	maxXmlElements = DEFAULT_XLSX_LOAD_LIMITS.maxXmlElements,
): void {
	try {
		checkXmlLimits(xml, label, { maxNodes: maxXmlElements });
	} catch (error) {
		throw workbookLimit(error);
	}
}
