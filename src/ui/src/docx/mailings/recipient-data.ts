import {
	mailMergeDelimitedChunks,
	mailMergeXlsx,
	selectMailMergeRecipients,
	type MailMergeData,
	type MailMergeRecord,
	type MailMergeValue,
} from 'ooxml-core/docx/mail-merge';

/**
 * What the editor accepts from one recipient source. Delimited text is read in 1 MiB slices and
 * never held whole; the records themselves stay in memory, so the caps bound that memory: at most
 * 32 Mi characters of values and 4 million values (for example 100,000 records of 40 fields).
 * An Excel workbook is a zip package that has to be read whole, so its file is capped lower.
 */
export const RECIPIENT_LIMITS = {
	delimitedBytes: 32 * 1024 * 1024,
	workbookBytes: 20 * 1024 * 1024,
	records: 100_000,
	fields: 255,
	cells: 4_000_000,
	characters: 32 * 1024 * 1024,
	valueLength: 16_384,
} as const;

const DATA_LIMITS = {
	maxFields: RECIPIENT_LIMITS.fields,
	maxRecords: RECIPIENT_LIMITS.records,
	maxCells: RECIPIENT_LIMITS.cells,
	maxCharacters: RECIPIENT_LIMITS.characters,
	maxValueLength: RECIPIENT_LIMITS.valueLength,
};
const CHUNK_BYTES = 1024 * 1024;

/** A file's text, decoded as UTF-8 one slice at a time. */
export async function* fileTextChunks(file: Blob, size = CHUNK_BYTES): AsyncGenerator<string> {
	const decoder = new TextDecoder();
	for (let start = 0; start < file.size; start += size) {
		const bytes = await file.slice(start, start + size).arrayBuffer();
		yield decoder.decode(bytes, { stream: true });
	}
	const rest = decoder.decode();
	if (rest) yield rest;
}

export async function readRecipientFile(file: File): Promise<MailMergeData> {
	const name = file.name.toLowerCase();
	if (name.endsWith('.xlsx')) {
		if (file.size > RECIPIENT_LIMITS.workbookBytes)
			throw new RangeError('Recipient workbook exceeds 20 MiB.');
		return checkedFields(
			await mailMergeXlsx(await file.arrayBuffer(), {
				dataLimits: DATA_LIMITS,
				limits: {
					maxInputBytes: RECIPIENT_LIMITS.workbookBytes,
					maxPartBytes: 128 * 1024 * 1024,
					maxTotalBytes: 192 * 1024 * 1024,
					maxParts: 512,
					maxXmlElements: 2_000_000,
				},
			}).read(),
		);
	}
	if (name.endsWith('.csv') || name.endsWith('.tsv')) {
		if (file.size > RECIPIENT_LIMITS.delimitedBytes)
			throw new RangeError('Recipient file exceeds 32 MiB.');
		return checkedFields(
			await mailMergeDelimitedChunks(() => fileTextChunks(file), {
				dataLimits: DATA_LIMITS,
				...(name.endsWith('.tsv') && { delimiter: '\t' }),
			}).read(),
		);
	}
	throw new TypeError('Unsupported recipient file type.');
}

function checkField(field: unknown): asserts field is string {
	if (
		typeof field !== 'string' ||
		!field.trim() ||
		field.length > 128 ||
		/[\u0000-\u001f\uE000\uE001]/.test(field)
	)
		throw new TypeError('Invalid recipient field name.');
}

/** Records read from a local file are built by the reader; only their field names need checks. */
function checkedFields(data: MailMergeData): MailMergeData {
	for (const field of data.fields ?? []) checkField(field);
	return data;
}

function recipientValue(value: unknown): value is MailMergeValue {
	return (
		value === null ||
		typeof value === 'string' ||
		typeof value === 'boolean' ||
		(typeof value === 'number' && Number.isFinite(value)) ||
		(value instanceof Date && Number.isFinite(value.getTime()))
	);
}

/** Copies the provider's records and places the same limits on all sources. */
export function snapshotRecipients(input: MailMergeData): MailMergeData {
	const { records, fields } = input;
	if (!Array.isArray(records) || records.length > DATA_LIMITS.maxRecords)
		throw new RangeError('Invalid recipient count.');
	const names = new Set<string>();
	let characters = 0;
	let cells = 0;
	const addField = (field: string) => {
		checkField(field);
		names.add(field);
		if (names.size > DATA_LIMITS.maxFields) throw new RangeError('Too many recipient fields.');
	};
	if (fields !== undefined) {
		if (!Array.isArray(fields) || fields.length > DATA_LIMITS.maxFields)
			throw new RangeError('Invalid recipient field count.');
		for (const field of fields) addField(field);
	}
	const copy = records.map((record: MailMergeRecord) => {
		if (!record || typeof record !== 'object' || Array.isArray(record))
			throw new TypeError('Invalid recipient.');
		const values: [string, MailMergeValue][] = [];
		for (const field in record) {
			if (!Object.hasOwn(record, field)) continue;
			addField(field);
			if (++cells > DATA_LIMITS.maxCells) throw new RangeError('Too many recipient values.');
			const value: unknown = record[field];
			if (!recipientValue(value)) throw new TypeError('Invalid recipient value.');
			if (typeof value === 'string') {
				characters += value.length;
				if (value.length > DATA_LIMITS.maxValueLength || characters > DATA_LIMITS.maxCharacters)
					throw new RangeError('Recipient text exceeds the size limit.');
			}
			values.push([field, value instanceof Date ? new Date(value.getTime()) : value]);
		}
		return Object.fromEntries(values);
	});
	return { records: copy, fields: [...names] };
}

export interface RecipientSelection {
	excluded: Set<MailMergeRecord>;
	filter: string;
	filterField: string;
	sortField: string;
	descending: boolean;
}

export function allRecipients(): RecipientSelection {
	return { excluded: new Set(), filter: '', filterField: '', sortField: '', descending: false };
}

export function recipientText(record: MailMergeRecord, field: string): string {
	const value = Object.hasOwn(record, field) ? record[field] : null;
	return value instanceof Date ? value.toISOString() : String(value ?? '');
}

export function recipientMatches(
	record: MailMergeRecord,
	fields: readonly string[],
	term: string,
	locale: string,
): boolean {
	return fields.some((field) =>
		recipientText(record, field).toLocaleLowerCase(locale).includes(term),
	);
}

export function selectedRecipients(
	data: MailMergeData,
	settings: RecipientSelection,
	locale: string,
	showExcluded = false,
): MailMergeData {
	const { excluded, filter, filterField, sortField, descending } = settings;
	const term = filter.trim().toLocaleLowerCase(locale);
	const fields = filterField ? [filterField] : (data.fields ?? []);
	const collator = new Intl.Collator(locale, { numeric: true });
	const direction = descending ? -1 : 1;
	return selectMailMergeRecipients(data, {
		filter: (record) =>
			(showExcluded || !excluded.has(record)) &&
			(!term || recipientMatches(record, fields, term, locale)),
		...(sortField && {
			sort: (left, right) =>
				direction *
				collator.compare(recipientText(left, sortField), recipientText(right, sortField)),
		}),
	});
}
