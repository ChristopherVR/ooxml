import {
	mailMergeDelimitedText,
	mailMergeXlsx,
	selectMailMergeRecipients,
	type MailMergeData,
	type MailMergeRecord,
	type MailMergeValue,
} from 'ooxml-core/docx/mail-merge';

const INPUT_BYTES = 5 * 1024 * 1024;
const DATA_LIMITS = { maxFields: 64, maxRecords: 1000 };
const MAX_DATA_CHARS = 2 * 1024 * 1024;

export async function readRecipientFile(file: File): Promise<MailMergeData> {
	if (file.size > INPUT_BYTES) throw new RangeError('Recipient file exceeds 5 MiB.');
	const name = file.name.toLowerCase();
	if (name.endsWith('.xlsx'))
		return snapshotRecipients(
			await mailMergeXlsx(await file.arrayBuffer(), {
				dataLimits: DATA_LIMITS,
				limits: {
					maxInputBytes: INPUT_BYTES,
					maxPartBytes: 12 * 1024 * 1024,
					maxTotalBytes: 24 * 1024 * 1024,
					maxParts: 512,
					maxXmlElements: 180_000,
				},
			}).read(),
		);
	if (name.endsWith('.csv') || name.endsWith('.tsv'))
		return snapshotRecipients(
			await mailMergeDelimitedText(await file.text(), {
				dataLimits: DATA_LIMITS,
				...(name.endsWith('.tsv') && { delimiter: '\t' }),
			}).read(),
		);
	throw new TypeError('Unsupported recipient file type.');
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
	const addField = (field: string) => {
		if (
			typeof field !== 'string' ||
			!field.trim() ||
			field.length > 128 ||
			/[\u0000-\u001f\uE000\uE001]/.test(field)
		)
			throw new TypeError('Invalid recipient field name.');
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
			const value: unknown = record[field];
			if (!recipientValue(value)) throw new TypeError('Invalid recipient value.');
			if (typeof value === 'string') {
				characters += value.length;
				if (value.length > 16_384 || characters > MAX_DATA_CHARS)
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
