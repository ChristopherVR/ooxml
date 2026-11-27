// The recipient readers that stream their source (delimited text read in slices, a workbook),
// loaded on their first read so the mail merge entry carries only the in-memory sources.
import { csvRowReader } from '../../xlsx/load/csv-stream';
import { isCellError } from '../../xlsx/model';
import type { ReadXlsxRowsOptions } from '../../xlsx/read/rows';
import type {
	MailMergeData,
	MailMergeDataLimits,
	MailMergeRecord,
	MailMergeValue,
} from './data-source';
import { readHeaders } from './record-headers';

/** The longest field name a streamed header may hold, whatever the value limit. */
const MAX_FIELD_NAME_LENGTH = 255;

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

/** The records of `mailMergeDelimitedChunks`. */
export async function readDelimitedChunks(
	chunks: () => AsyncIterable<string> | Iterable<string>,
	options: { delimiter?: ',' | '\t' | ';'; dataLimits?: MailMergeDataLimits } = {},
): Promise<MailMergeData> {
	const limits = options.dataLimits;
	const reader = csvRowReader({
		...(options.delimiter && { delimiter: options.delimiter }),
		...(limits && {
			maxRows: limits.maxRecords + 1,
			maxColumns: limits.maxFields,
			...(limits.maxValueLength !== undefined && {
				maxFieldLength: limits.maxValueLength,
				// The value limit is for recipients' values, not the field names of row 1.
				maxHeaderFieldLength: Math.max(limits.maxValueLength, MAX_FIELD_NAME_LENGTH),
			}),
		}),
	});
	const builder = new RecordBuilder(limits);
	for await (const chunk of chunks()) for (const row of reader.push(chunk)) builder.add(row);
	for (const row of reader.end()) builder.add(row);
	return { fields: builder.fieldNames, records: builder.records };
}

/** The records of `mailMergeXlsx`. */
export async function readXlsxRecords(
	bytes: Uint8Array | ArrayBuffer,
	options?: ReadXlsxRowsOptions & { dataLimits?: MailMergeDataLimits },
): Promise<MailMergeData> {
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
}
