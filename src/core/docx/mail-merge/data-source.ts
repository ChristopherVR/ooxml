import { parseCsv } from '../../xlsx/load/csv';
import { readHeaders } from './record-headers';
import type { ReadXlsxRowsOptions } from '../../xlsx/read/rows';

export type MailMergeValue = string | number | boolean | Date | null;
export type MailMergeRecord = Readonly<Record<string, MailMergeValue>>;

/** An in-memory recipient list. Fields can be supplied even when the source is empty. */
export interface MailMergeData {
	readonly records: readonly MailMergeRecord[];
	readonly fields?: readonly string[];
}

/** A host supplies records and owns any access to external services. */
export interface MailMergeDataSource {
	read(): MailMergeData | Promise<MailMergeData>;
}

export interface MailMergeDataLimits {
	maxFields: number;
	maxRecords: number;
	/** Text held across all values, in UTF-16 code units. */
	maxCharacters?: number;
	/** The longest single value, in UTF-16 code units. */
	maxValueLength?: number;
	/** Records times fields: bounds the memory the record objects take. */
	maxCells?: number;
}

export function mailMergeRecords(data: MailMergeData): MailMergeDataSource {
	return { read: () => data };
}

/** CSV, TSV, or auto-detected delimited text. Values stay strings, including leading zeroes. */
export function mailMergeDelimitedText(
	text: string,
	options: { delimiter?: ',' | '\t' | ';'; dataLimits?: MailMergeDataLimits } = {},
): MailMergeDataSource {
	return {
		read() {
			const limits = options.dataLimits;
			const [header = [], ...rows] = parseCsv(text, {
				...options,
				...(limits && { maxRows: limits.maxRecords + 1, maxColumns: limits.maxFields }),
			});
			const fields = readHeaders(header.entries());
			const records = rows.map((row) =>
				Object.fromEntries([...fields].map(([col, field]) => [field, row[col] ?? ''])),
			);
			return { fields: [...fields.values()], records };
		},
	};
}

/**
 * CSV or TSV text read one chunk at a time (for example a `File` read in slices), so the whole
 * file is never one string: each completed row becomes a record and the text it came from can be
 * collected. The same rules as `mailMergeDelimitedText` apply.
 */
export function mailMergeDelimitedChunks(
	chunks: () => AsyncIterable<string> | Iterable<string>,
	options: { delimiter?: ',' | '	' | ';'; dataLimits?: MailMergeDataLimits } = {},
): MailMergeDataSource {
	return {
		async read() {
			// The streaming reader loads with the first read; in-memory sources never need it.
			return (await import('./data-source-streams')).readDelimitedChunks(chunks, options);
		},
	};
}

/** Reads row 1 as field names and later rows as recipients. Formula cells use their saved value. */
export function mailMergeXlsx(
	bytes: Uint8Array | ArrayBuffer,
	options?: ReadXlsxRowsOptions & { dataLimits?: MailMergeDataLimits },
): MailMergeDataSource {
	return {
		async read() {
			return (await import('./data-source-streams')).readXlsxRecords(bytes, options);
		},
	};
}

export interface MailMergeRecipientOptions {
	filter?: (record: MailMergeRecord, sourceIndex: number) => boolean;
	sort?: (left: MailMergeRecord, right: MailMergeRecord) => number;
}

/** Filters by source position (zero-based), then sorts stably, leaving the input list untouched. */
export function selectMailMergeRecipients(
	data: MailMergeData,
	{ filter, sort }: MailMergeRecipientOptions = {},
): MailMergeData {
	const records = filter ? data.records.filter(filter) : [...data.records];
	if (sort) records.sort(sort);
	return { records, fields: data.fields ?? [...new Set(data.records.flatMap(Object.keys))] };
}
