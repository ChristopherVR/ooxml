import type { ParsedField } from '../field-format';
import { fieldTokens } from '../field-instruction';

export const addressSwitches = { c: 'arg', d: 'flag', e: 'arg', f: 'arg', l: 'arg' } as const;
export const greetingSwitches = { e: 'arg', f: 'arg', l: 'arg' } as const;

const placeholders = [
	['_TITLE0_', '/t0', 'CourtesyTitle'],
	['_NICK0_', '/n0', 'Nickname'],
	['_FIRST0_', '/f0', 'FirstName'],
	['_MIDDLE0_', '/m0', 'MiddleName'],
	['_LAST0_', '/l0', 'LastName'],
	['_SUFFIX0_', '/s0', 'Suffix'],
	['_TITLE1_', '/t1', 'SpouseCourtesyTitle'],
	['_FIRST1_', '/f1', 'SpouseFirstName'],
	['_MIDDLE1_', '/m1', 'SpouseMiddleName'],
	['_LAST1_', '/l1', 'SpouseLastName'],
	['_SUFFIX1_', '/s1', 'SpouseNickname'],
	['_COMPANY_', '/c', 'Company'],
	['_STREET1_', '/r', 'Address1'],
	['_STREET2_', '/e', 'Address2'],
	['_CITY_', '/i', 'City'],
	['_STATE_', '/a', 'State'],
	['_POSTAL_', '/z', 'PostalCode'],
	['_COUNTRY_', '/d', 'CountryorRegion'],
] as const;
const columns = new Map<string, string>(
	placeholders.flatMap(([name, alias, column]) => [
		[name, column],
		[alias, column],
	]),
);
const placeholder = /_[A-Z]+\d?_|\/[a-z]\d?/g;
const addressName = '<<_TITLE0_ >><<_FIRST0_>><< _LAST0_>><< _SUFFIX0_>>';
const firstLast = '<<_FIRST0_>><< _LAST0_>><< _SUFFIX0_>>';
const titleLast = '<<_TITLE0_>><< _LAST0_>>';
export const defaultGreetingFallback = 'Dear Sir or Madam,';
const internationalLocality = '<<_POSTAL_ >><<_CITY_>><< _STATE_>>';
const domesticLocality = '<<_CITY_>><<, _STATE_>><< _POSTAL_>>';
const postalCountry =
	/^(de(?:utschland)?|germany|fr(?:ance)?|es(?:paña)?|spain|it(?:alia|aly)?|nl|netherlands|be(?:lgium)?|at|austria|ch|switzerland)$/i;
const postalLanguage =
	/^(?:1031|1034|1036|1040|1043|3082|0407|040a|040c|0410|0413|de|fr|es|it|nl)(?:-|$)/i;

const localeIds = new Map([
	['en', 1033],
	['en-gb', 2057],
	['en-au', 3081],
	['de', 1031],
	['es', 3082],
	['fr', 1036],
	['it', 1040],
	['nl', 1043],
	['zh', 2052],
]);

/** Word writes the field's language as an unquoted decimal Windows LCID. */
export function mailingLanguageId(language = ''): number {
	const tag = language.trim().toLowerCase();
	const first = tag.split('-')[0] ?? '';
	if (/^(?:0x[\da-f]{4}|0[\da-f]{3})$/.test(tag)) return Number.parseInt(tag, 16);
	if (/^[1-9]\d{0,4}$/.test(tag) && Number(tag) <= 65535) return Number(tag);
	return localeIds.get(tag) ?? localeIds.get(first) ?? 1033;
}

export interface AddressBlockOptions {
	name?: 'full' | 'first' | 'last' | 'none';
	company?: boolean;
	street?: boolean;
	country?: '0' | '1' | '2';
	excludeCountry?: string;
	formatByCountry?: boolean;
	language?: string;
}

export function addressTemplate(options: AddressBlockOptions = {}): string {
	const name = {
		full: addressName,
		first: '<<_FIRST0_>>',
		last: '<<_TITLE0_ >><<_LAST0_>>',
		none: '',
	};
	return [
		name[options.name ?? 'full'],
		'\r',
		options.company !== false ? '<<_COMPANY_\r>>' : '',
		options.street !== false
			? `<<_STREET1_\r>><<_STREET2_\r>>${domesticLocality}<<\r_COUNTRY_>>`
			: '',
	].join('');
}

export interface GreetingLineOptions {
	name?: 'full' | 'first' | 'last';
	before?: string;
	after?: string;
	fallback?: string;
	language?: string;
}

export function greetingTemplate(options: GreetingLineOptions = {}): string {
	const name = { full: firstLast, first: '<<_FIRST0_>>', last: titleLast };
	const before = (options.before ?? 'Dear').replace(/[<>]/g, '').trimEnd();
	const after = (options.after ?? ',').replace(/[<>]/g, '');
	return `<<_BEFORE_ ${before ? `${before} ` : ''}>>${name[options.name ?? 'last']}\n<<_AFTER_ ${after}>>`;
}

export function renderMailingField(
	field: ParsedField,
	instruction: string,
	lookup: (name: string) => string,
	resolve: (arg: string) => string,
): string | undefined {
	if (field.args.length || [...field.switches].some(([key, arg]) => arg === true && key !== 'd'))
		return undefined;
	const arg = (name: string) => {
		const value = field.switches.get(name);
		return typeof value === 'string' ? resolve(value) : undefined;
	};
	const greeting = field.name === 'GREETINGLINE';
	const language = arg('l') ?? '';
	const country = lookup('CountryorRegion');
	const countryMode = arg('c') ?? '1';
	if (!['0', '1', '2'].includes(countryMode)) return undefined;
	const excluded: string[] = [];
	if (countryMode === '2') {
		const tokens = fieldTokens(instruction);
		for (const [index, token] of tokens.entries()) {
			const next = tokens[index + 1];
			if (token.kind === 'switch' && token.value === 'e' && next?.kind === 'text')
				excluded.push(resolve(next.value).trim().toLowerCase());
		}
	}
	const addCountry = countryMode !== '0' && !excluded.includes(country.toLowerCase());
	let template = arg('f') ?? (greeting ? greetingTemplate() : addressTemplate());
	if (template.length > 4096) return undefined;
	template = template.replace(/\r\n?/g, '\n');
	if (
		!greeting &&
		(field.switches.has('d') && country
			? postalCountry.test(country)
			: postalLanguage.test(language))
	)
		template = template.replace(domesticLocality, internationalLocality);
	let missing = false;
	let before = '';
	let after = '';
	let substitutions = 0;
	const replace = (part: string) => {
		let empty = false;
		const value = part.replace(placeholder, (name) => {
			if (++substitutions > 64) return '';
			const column = columns.get(name);
			const result = column && (column !== 'CountryorRegion' || addCountry) ? lookup(column) : '';
			empty ||= !result;
			missing ||= !result && name !== '_SUFFIX0_';
			return result.slice(0, 1024);
		});
		return empty ? '' : value;
	};
	const pieces: string[] = [];
	let separator = '';
	// Between groups, whitespace separates outputs. Inside a group, it is part of the output.
	const append = (value: string, grouped: boolean) => {
		if (!value) return;
		if (!grouped && !value.trim()) {
			if (pieces.length) separator += value;
		} else {
			pieces.push(separator, value);
			separator = '';
		}
	};
	let cursor = 0;
	for (const group of template.matchAll(/<<([^<>]*)>>/g)) {
		append(replace(template.slice(cursor, group.index)), false);
		const item = group[1] ?? '';
		if (greeting && /^(?:_BEFORE_|\/g)(?: |$)/.test(item))
			before = item.replace(/^(?:_BEFORE_|\/g) ?/, '');
		else if (greeting && /^(?:_AFTER_|\/x)(?: |$)/.test(item))
			after = item.replace(/^(?:_AFTER_|\/x) ?/, '');
		else append(replace(item), true);
		cursor = group.index + group[0].length;
	}
	append(replace(template.slice(cursor)), false);
	let text = pieces.join('');
	if (substitutions > 64) return undefined;
	if (!greeting) return text;
	text = text.trim();
	if (!text || missing) {
		const name = template.replace(/<<(?:_BEFORE_|_AFTER_|\/g|\/x)(?: [^<>]*)?>>/g, '').trim();
		// Word 16 can fall back to the first name, with the last name when it is present.
		const first =
			name === titleLast || name === firstLast ? lookup('FirstName').slice(0, 1024) : '';
		if (!first) return arg('e') ?? '';
		text = `${first} ${lookup('LastName').slice(0, 1024)}`.trimEnd();
	}
	return `${before && !/\s$/.test(before) ? `${before} ` : before}${text}${after}`;
}
