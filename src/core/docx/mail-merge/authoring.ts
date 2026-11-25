import type { TextRun } from '../model';
import type { MailMergeFieldMapping } from './settings';

/** The address columns offered by Match Fields and the composed field commands. */
export const mailMergeAddressFields = [
	'FirstName',
	'LastName',
	'Company',
	'Address1',
	'Address2',
	'City',
	'State',
	'PostalCode',
	'CountryorRegion',
] as const;
type AddressField = (typeof mailMergeAddressFields)[number];

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

export function mailMergeFieldMap(
	fields: readonly string[],
	mappings: readonly MailMergeFieldMapping[] = [],
): ReadonlyMap<string, string> {
	const map = new Map(fields.map((field) => [field.toLowerCase(), field]));
	const column = (item: MailMergeFieldMapping) => item.name || fields[item.column ?? -1];
	for (const [index, item] of mappings.entries()) {
		const name = ODSO_NAMES[index];
		const value = column(item);
		if (name && value && item.type !== 'null') map.set(name.toLowerCase(), value);
	}
	for (const item of mappings) {
		if (!item.mappedName) continue;
		const value = column(item);
		if (item.type === 'null') map.delete(item.mappedName.toLowerCase());
		else if (value) map.set(item.mappedName.toLowerCase(), value);
	}
	return map;
}

export function matchMailMergeField(
	name: string,
	fields: readonly string[],
	mappings: readonly MailMergeFieldMapping[] = [],
): string {
	const candidate = mailMergeFieldMap(fields, mappings).get(name.toLowerCase());
	return candidate && fields.includes(candidate) ? candidate : '';
}

export function setMailMergeFieldMapping(
	mappings: MailMergeFieldMapping[],
	name: string,
	fields: readonly string[],
	value: string,
): void {
	const key = name.toLowerCase();
	let explicit = mappings.length - 1;
	while (explicit >= 0 && mappings[explicit]?.mappedName?.toLowerCase() !== key) explicit--;
	const position = ODSO_NAMES.findIndex((item) => item === name);
	const occupied = mappings[position]?.mappedName;
	const index =
		explicit >= 0 ? explicit : occupied ? Math.max(mappings.length, ODSO_NAMES.length) : position;
	if (index < 0) return;
	while (mappings.length <= index) mappings.push({ type: 'null' });
	mappings[index] = value
		? { type: 'dbColumn', name: value, mappedName: name, column: fields.indexOf(value) }
		: { type: 'null', mappedName: name };
}

export function composeMailMergeFields(
	kind: 'address' | 'greeting',
	fields: readonly string[],
	mappings: readonly MailMergeFieldMapping[],
	greeting = '',
): TextRun[] {
	const name = ['FirstName', 'LastName'] as const;
	const lines: readonly (readonly AddressField[])[] =
		kind === 'greeting'
			? [name]
			: [
					name,
					['Company'],
					['Address1'],
					['Address2'],
					['City', 'State', 'PostalCode'],
					['CountryorRegion'],
				];
	const matched = lines
		.map((line) => line.filter((field) => matchMailMergeField(field, fields, mappings)))
		.filter((line) => line.length);
	const runs = matched.flatMap((line, lineIndex) => [
		...(lineIndex ? [{ text: '\n' }] : []),
		...line.map((field, index) =>
			mailMergeField(field, { mapped: true, ...(index && { before: ' ' }) }),
		),
	]);
	return kind === 'greeting' && runs.length
		? [{ text: greeting ? `${greeting} ` : '' }, ...runs, { text: ',' }]
		: runs;
}

export type MailMergeComparison = '=' | '<>' | '<' | '<=' | '>' | '>=';
export const mailMergeComparisons: readonly MailMergeComparison[] = [
	'=',
	'<>',
	'<',
	'<=',
	'>',
	'>=',
];

export type MailMergeRule =
	| { kind: 'NEXT' | 'MERGEREC' | 'MERGESEQ' }
	| {
			kind: 'IF' | 'NEXTIF' | 'SKIPIF';
			field: string;
			operator: MailMergeComparison;
			value: string;
			yes: string;
			no: string;
	  }
	| { kind: 'ASK' | 'FILLIN'; bookmark: string; prompt: string; answer: string }
	| { kind: 'REF'; bookmark: string };

// The field parser unescapes these two characters inside quoted arguments.
const quote = (text: string) => `"${text.replace(/["\\]/g, '\\$&')}"`;
const code = (fieldCode: string): TextRun => ({ text: '', fieldCode });

function complex(instruction: TextRun[], label: string): TextRun[] {
	return [
		{ text: '', fieldChar: 'begin' },
		...instruction,
		{ text: '', fieldChar: 'separate' },
		{
			text: `«${label}»`,
			field: {
				instr: instruction
					.map((run) => run.fieldCode ?? '')
					.join('')
					.trim(),
			},
		},
		{ text: '', fieldChar: 'end' },
	];
}

export function mailMergeField(
	name: string,
	options: { mapped?: boolean; before?: string; after?: string } = {},
): TextRun {
	const switches = [
		options.mapped ? ' \\m' : '',
		options.before ? ` \\b ${quote(options.before)}` : '',
		options.after ? ` \\f ${quote(options.after)}` : '',
	].join('');
	return {
		text: `«${name}»`,
		field: { simple: true, instr: `MERGEFIELD ${quote(name)}${switches}` },
	};
}

export function mailMergeRule(rule: MailMergeRule): TextRun[] {
	const { kind } = rule;
	if (kind === 'IF' || kind === 'NEXTIF' || kind === 'SKIPIF') {
		const operand = mailMergeField(rule.field);
		const branches = kind === 'IF' ? ` ${quote(rule.yes)} ${quote(rule.no)}` : '';
		return complex(
			[
				code(` ${kind} `),
				...complex([code(` ${operand.field?.instr} `)], rule.field),
				code(` ${rule.operator} ${quote(rule.value)}${branches} `),
			],
			kind,
		);
	}
	let args = '';
	if (kind === 'ASK' || kind === 'FILLIN')
		args = ` ${kind === 'ASK' ? `${quote(rule.bookmark)} ` : ''}${quote(rule.prompt)} \\d ${quote(rule.answer)}`;
	if (kind === 'REF') args = ` ${quote(rule.bookmark)}`;
	return complex([code(` ${kind}${args} `)], kind);
}
