import { formatDatePicture } from '../../text/date-picture';
import { compareFieldValues, evaluateIfField } from '../conditional-field';
import { applyFieldFormat, parseFieldInstruction, type FieldFormat } from '../field-format';
import { fieldName } from '../field-runs';
import type { MailMergeData, MailMergeValue } from './data-source';
import { formatMergeNumber } from './number-picture';
import type { MailMergeFieldMapping } from './settings';

const MERGE_SWITCHES = { b: 'arg', f: 'arg', m: 'flag', v: 'flag' } as const;
const PROMPT_SWITCHES = { d: 'arg', o: 'flag' } as const;
const MERGE_FIELDS = new Set([
	'MERGEFIELD',
	'NEXT',
	'NEXTIF',
	'SKIPIF',
	'MERGEREC',
	'MERGESEQ',
	'ASK',
	'FILLIN',
]);
const UNSUPPORTED_FIELDS = new Set(['ADDRESSBLOCK', 'GREETINGLINE', 'MERGEBARCODE']);

// [MS-OI29500] 2.1.369 defines the positional ODSO names Word writes, including empty slots.
const ODSO_NAMES = [
	'Unique',
	'CourtesyTitle',
	'FirstName',
	'MiddleName',
	'LastName',
	'Suffix',
	'Nickname',
	'JobTitle',
	'Company',
	'Address1',
	'Address2',
	'City',
	'State',
	'PostalCode',
	'CountryorRegion',
	'BusinessPhone',
	'BusinessFax',
	'HomePhone',
	'HomeFax',
	'EmailAddress',
	'WebPage',
	'SpouseCourtesyTitle',
	'SpouseFirstName',
	'SpouseMiddleName',
	'SpouseLastName',
	'SpouseNickname',
	'RubyFirstName',
	'RubyLastName',
	'Address3',
	'Department',
];

export type MailMergeIssueCode =
	| 'no-recipients'
	| 'unknown-field'
	| 'invalid-instruction'
	| 'unsupported-field'
	| 'invalid-record-rule-location';

type FieldIssueCode = Exclude<MailMergeIssueCode, 'no-recipients'>;

export type MailMergeIssue =
	| { code: 'no-recipients' }
	| {
			code: FieldIssueCode;
			/** Position in the filtered and sorted list, one-based. */
			recordNumber: number;
			blockId: string;
			field: string;
	  };

export interface MailMergeOptions {
	locale?: string;
	/** Overrides the mappings read from the template. */
	fieldMappings?: readonly MailMergeFieldMapping[];
}

interface FieldResult {
	text: string;
	charFormat?: true;
}

function textOf(value: MailMergeValue): string {
	if (value instanceof Date) return Number.isFinite(value.getTime()) ? value.toISOString() : '';
	if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
	return String(value ?? '');
}

function dateOf(value: MailMergeValue): Date | undefined {
	if (value instanceof Date) return value;
	if (typeof value !== 'string') return undefined;
	const text = value.trim();
	const dateOnly = /^\d{4}-\d{2}-\d{2}$/.test(text);
	const timeOnly = /^\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?$/.test(text);
	const isoTime = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(Z|[+-]\d{2}:\d{2})?$/i.exec(
		text,
	);
	if (dateOnly) return new Date(`${text}T00:00:00Z`);
	if (timeOnly) return new Date(`1970-01-01T${text}Z`);
	if (isoTime) return new Date(isoTime[1] ? text : `${text}Z`);
	const us = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(text);
	if (us)
		return new Date(`${us[3]}-${us[1]?.padStart(2, '0')}-${us[2]?.padStart(2, '0')}T00:00:00Z`);
	return undefined;
}

function formatMergeValue(value: MailMergeValue, format: FieldFormat, locale: string): string {
	let text = textOf(value);
	if (format.datePicture !== undefined) {
		const date = dateOf(value);
		if (date && Number.isFinite(date.getTime()))
			text = formatDatePicture(date, format.datePicture, { locale, utc: true });
	} else if (
		format.numberPicture !== undefined &&
		/^[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?$/i.test(text.trim())
	) {
		const number = Number(text);
		if (Number.isFinite(number)) text = formatMergeNumber(number, format.numberPicture);
	}
	return applyFieldFormat(text, format);
}

/** Character substitution used for vertical East Asian address text. */
function verticalText(text: string): string {
	return text.replace(/[0-9A-Za-z\-－ー]/g, (char) => {
		if (/\d/.test(char)) return '〇一二三四五六七八九'.charAt(Number(char));
		if (/^[A-Za-z]$/.test(char)) return String.fromCharCode(char.charCodeAt(0) + 0xfee0);
		return '︱';
	});
}

export function isMergeField(name: string): boolean {
	return MERGE_FIELDS.has(name) || UNSUPPORTED_FIELDS.has(name);
}

/** State belonging to one output copy of the template. */
export class MergeFields {
	index: number;
	skipped = false;
	readonly issues: MailMergeIssue[] = [];
	private readonly fields: ReadonlySet<string>;
	private readonly mappings = new Map<string, string>();
	private readonly answers = new Map<string, string>();

	constructor(
		private readonly data: MailMergeData,
		start: number,
		private readonly sequence: number,
		private readonly options: MailMergeOptions,
	) {
		this.index = start;
		const fields = data.fields ?? [...new Set(data.records.flatMap(Object.keys))];
		this.fields = new Set(fields);
		const mappings = options.fieldMappings ?? [];
		mappings.forEach((mapping, index) => {
			if (mapping.type === 'null') return;
			const column =
				mapping.name || (mapping.column !== undefined ? fields[mapping.column] : undefined);
			const name = ODSO_NAMES[index];
			if (name && column) this.mappings.set(name.toLowerCase(), column);
		});
		for (const mapping of mappings) {
			if (!mapping.mappedName || mapping.type === 'null') continue;
			const column =
				mapping.name || (mapping.column !== undefined ? fields[mapping.column] : undefined);
			if (column) this.mappings.set(mapping.mappedName.toLowerCase(), column);
		}
	}

	issue(code: FieldIssueCode, blockId: string, field: string): undefined {
		this.issues.push({ code, recordNumber: this.index + 1, blockId, field });
		return undefined;
	}

	evaluate(input: {
		instruction: string;
		blockId: string;
		cached: string;
		allowRecordRules: boolean;
		nestedMerge: boolean;
		resolve: (arg: string) => string;
	}): FieldResult | undefined {
		const { instruction, blockId, cached, allowRecordRules, nestedMerge, resolve } = input;
		const name = fieldName(instruction);
		const issue = (code: FieldIssueCode, field = name) => this.issue(code, blockId, field);
		if (UNSUPPORTED_FIELDS.has(name) || (nestedMerge && !MERGE_FIELDS.has(name) && name !== 'IF'))
			return issue('unsupported-field');
		const isPrompt = name === 'ASK' || name === 'FILLIN';
		const isRef = name === 'REF';
		if (!MERGE_FIELDS.has(name) && !isRef && !(name === 'IF' && nestedMerge)) return undefined;
		const own = name === 'MERGEFIELD' ? MERGE_SWITCHES : isPrompt ? PROMPT_SWITCHES : {};
		const field = parseFieldInstruction(instruction, own, { numberPictures: true });
		if (!field) return isRef ? undefined : issue('invalid-instruction');
		const [first, second, third] = field.args;
		const format = (value: MailMergeValue) =>
			formatMergeValue(value, field.format, this.options.locale ?? 'en-US');
		const result = (text: string): FieldResult => ({
			text,
			...(field.format.charFormat && { charFormat: true }),
		});
		if (name === 'MERGEFIELD') {
			const before = field.switches.get('b');
			const after = field.switches.get('f');
			if (!first || field.args.length !== 1 || before === true || after === true)
				return issue('invalid-instruction');
			const lookup = resolve(first);
			const column = field.switches.has('m')
				? (this.mappings.get(lookup.toLowerCase()) ?? lookup)
				: lookup;
			if (!this.fields.has(column)) issue('unknown-field', column);
			const record = this.data.records[this.index];
			const value = record && Object.hasOwn(record, column) ? (record[column] ?? null) : null;
			if (!textOf(value)) return result('');
			const text = format(value);
			return result(
				`${before === undefined ? '' : resolve(before)}${field.switches.has('v') ? verticalText(text) : text}${after === undefined ? '' : resolve(after)}`,
			);
		}
		if (isPrompt) {
			const fallback = field.switches.get('d');
			if (
				first === undefined ||
				field.args.length !== (name === 'ASK' ? 2 : 1) ||
				fallback === true
			)
				return issue('invalid-instruction');
			const answer = fallback === undefined ? cached : resolve(fallback);
			if (name === 'FILLIN') return result(format(answer));
			this.answers.set(resolve(first).toLowerCase(), answer);
			return result('');
		}
		if (isRef) {
			const answer =
				first && field.args.length === 1 ? this.answers.get(first.toLowerCase()) : undefined;
			return answer === undefined ? undefined : result(format(answer));
		}
		if (name === 'IF') {
			const text = evaluateIfField(field, resolve);
			return text === undefined ? issue('invalid-instruction') : result(text);
		}
		if (name === 'MERGEREC' || name === 'MERGESEQ') {
			if (field.args.length) return issue('invalid-instruction');
			if (name === 'MERGEREC' && !this.data.records[this.index]) return result('');
			return result(format(name === 'MERGEREC' ? this.index + 1 : this.sequence));
		}
		if (!allowRecordRules) return issue('invalid-record-rule-location');
		const matches =
			name === 'NEXT'
				? field.args.length
					? undefined
					: true
				: first !== undefined &&
					  second !== undefined &&
					  third !== undefined &&
					  field.args.length === 3
					? compareFieldValues(resolve(first), second, resolve(third))
					: undefined;
		if (matches === undefined) return issue('invalid-instruction');
		if (matches && this.index < this.data.records.length) {
			if (name === 'SKIPIF') this.skipped = true;
			else this.index++;
		}
		return result('');
	}
}
