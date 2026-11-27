import { parseCsv } from '../../xlsx/load/csv';
import { csvRowReader } from '../../xlsx/load/csv-stream';
import { isCellError } from '../../xlsx/model';
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

function readHeaders(cells: Iterable<readonly [number, string]>): Map<number, string> {
	const fields = new Map<number, string>();
	const seen = new Set<string>();
	for (const [column, raw] of cells) {
		const field = raw.trim();
		if (!field) continue;
		if (seen.has(field)) throw new Error(`Duplicate mail merge field: ${field}`);
		seen.add(field);
		fields.set(column, field);
	}
	return fields;
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

/** Builds records row by row, enforcing `limits` as each row arrives. */
class RecordBuilder {
	private fields: Map<number, string> | undefined;
	readonly records: MailMergeRecord[] = [];
	private characters = 0;

	constructor(private readonly limits: MailMergeDataLimits | undefined) {}

	get fieldNames(): string[] {
		return [...(this.fields?.values() ?? [])];
	}

	add(row: readonly string[]): void {
		if (!this.fields) {
			this.fields = readHeaders(row.entries());
			return;
		}
		const limits = this.limits;
		if (limits) {
			if (this.records.length >= limits.maxRecords)
				throw new RangeError('Too many mail merge rows.');
			if (
				limits.maxCells !== undefined &&
				(this.records.length + 1) * this.fields.size > limits.maxCells
			)
				throw new RangeError('Too many mail merge values.');
		}
		const values: [string, string][] = [];
		for (const [col, field] of this.fields) {
			const value = row[col] ?? '';
			if (limits?.maxValueLength !== undefined && value.length > limits.maxValueLength)
				throw new RangeError('A mail merge value is too long.');
			this.characters += value.length;
			if (limits?.maxCharacters !== undefined && this.characters > limits.maxCharacters)
				throw new RangeError('The mail merge data is too large.');
			values.push([field, value]);
		}
		this.records.push(Object.fromEntries(values));
	}
}

/**
 * CSV or TSV text read one chunk at a time (for example a `File` read in slices), so the whole
 * file is never one string: each completed row becomes a record and the text it came from can be
 * collected. The same rules as `mailMergeDelimitedText` apply.
 */
export function mailMergeDelimitedChunks(
	chunks: () => AsyncIterable<string> | Iterable<string>,
	options: { delimiter?: ',' | '\t' | ';'; dataLimits?: MailMergeDataLimits } = {},
): MailMergeDataSource {
	return {
		async read() {
			const limits = options.dataLimits;
			const reader = csvRowReader({
				...(options.delimiter && { delimiter: options.delimiter }),
				...(limits && {
					maxRows: limits.maxRecords + 1,
					maxColumns: limits.maxFields,
					...(limits.maxValueLength !== undefined && { maxFieldLength: limits.maxValueLength }),
				}),
			});
			const builder = new RecordBuilder(limits);
			for await (const chunk of chunks()) for (const row of reader.push(chunk)) builder.add(row);
			for (const row of reader.end()) builder.add(row);
			return { fields: builder.fieldNames, records: builder.records };
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
			const { readXlsxRows } = await import('../../xlsx/read/rows');
			let fields = new Map<number, string>();
			const records: MailMergeRecord[] = [];
			const limits = options?.dataLimits;
			let characters = 0;
			const counted = (value: MailMergeValue) => {
				if (typeof value === 'string' && limits) {
					if (limits.maxValueLength !== undefined && value.length > limits.maxValueLength)
						throw new RangeError('A mail merge value is too long.');
					characters += value.length;
					if (limits.maxCharacters !== undefined && characters > limits.maxCharacters)
						throw new RangeError('The mail merge data is too large.');
				}
				return value;
			};
			for await (const row of readXlsxRows(bytes, options)) {
				if (row.row === 0) {
					if (options?.dataLimits && row.cells.length > options.dataLimits.maxFields)
						throw new RangeError('Too many mail merge columns.');
					fields = readHeaders(row.cells.map(({ col, value }) => [col, String(value ?? '')]));
					continue;
				}
				if (!fields.size) throw new Error('The worksheet needs field names in row 1.');
				if (limits && records.length >= limits.maxRecords)
					throw new RangeError('Too many mail merge rows.');
				if (limits?.maxCells !== undefined && (records.length + 1) * fields.size > limits.maxCells)
					throw new RangeError('Too many mail merge values.');
				const cells = new Map(row.cells.map((cell) => [cell.col, cell]));
				records.push(
					Object.fromEntries(
						[...fields].map(([col, field]) => {
							const cell = cells.get(col);
							const value = cell?.value;
							return [
								field,
								counted(cell?.iso ?? (isCellError(value) ? value.error : (value ?? null))),
							];
						}),
					),
				);
			}
			return { fields: [...fields.values()], records };
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
