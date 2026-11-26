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
const addressName = '<<_FIRST0_>><<_LAST0_>>';
const greetingName = addressName;
const internationalLocality = '<<_POSTAL_>><<_CITY_>><<_STATE_>>';
const domesticLocality = '<<_CITY_>><<, _STATE_>><<_POSTAL_>>';
const postalCountry =
	/^(de(?:utschland)?|germany|fr(?:ance)?|es(?:paña)?|spain|it(?:alia|aly)?|nl|netherlands|be(?:lgium)?|at|austria|ch|switzerland)$/i;
const postalLanguage =
	/^(?:1031|1034|1036|1040|1043|0407|040a|040c|0410|0413|de|fr|es|it|nl)(?:-|$)/i;

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
		last: '<<_TITLE0_>><<_LAST0_>>',
		none: '',
	};
	return [
		name[options.name ?? 'full'],
		options.company !== false ? '<<_COMPANY_>>' : '',
		...(options.street !== false
			? ['<<_STREET1_>>', '<<_STREET2_>>', domesticLocality, '<<_COUNTRY_>>']
			: []),
	]
		.filter(Boolean)
		.join('\n');
}

export interface GreetingLineOptions {
	name?: 'full' | 'first' | 'last';
	before?: string;
	after?: string;
	fallback?: string;
	language?: string;
}

export function greetingTemplate(options: GreetingLineOptions = {}): string {
	const name = { full: greetingName, first: '<<_FIRST0_>>', last: '<<_TITLE0_>><<_LAST0_>>' };
	return `<<_BEFORE_ ${(options.before ?? 'Dear').replace(/[<>]/g, '')}>>${name[options.name ?? 'full']}<<_AFTER_ ${(options.after ?? ',').replace(/[<>]/g, '')}>>`;
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
			return result.slice(0, 1024);
		});
		missing ||= empty;
		return empty ? '' : value;
	};
	const lines = template
		.split(/\r\n?|\n/)
		.map((line) => {
			const pieces: string[] = [];
			let cursor = 0;
			for (const group of line.matchAll(/<<([^<>]*)>>/g)) {
				pieces.push(replace(line.slice(cursor, group.index)));
				const item = group[1] ?? '';
				if (greeting && /^(?:_BEFORE_|\/g)(?: |$)/.test(item))
					before = item.replace(/^(?:_BEFORE_|\/g) ?/, '');
				else if (greeting && /^(?:_AFTER_|\/x)(?: |$)/.test(item))
					after = item.replace(/^(?:_AFTER_|\/x) ?/, '');
				else {
					const value = replace(item);
					if (value && pieces.some((piece) => piece.trim()) && !/^[\s,.;:]/.test(value))
						pieces.push(' ');
					pieces.push(value);
				}
				cursor = group.index + group[0].length;
			}
			pieces.push(replace(line.slice(cursor)));
			return pieces.join('').trim();
		})
		.filter(Boolean);
	const text = lines.join('\n');
	if (substitutions > 64) return undefined;
	if (!greeting) return text;
	if (!text || missing) return arg('e') ?? '';
	return `${before ? `${before} ` : ''}${text}${after}`;
}
