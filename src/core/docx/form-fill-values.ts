// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Checking a value for a form field against its constraints and writing it as the field shows it:
// a date with the field's date picture, a number with its number picture, a legacy text field's
// text case. A value that does not fit is refused with the reason, never guessed at.
import { formatWordDate } from './field-date';
import { applyFieldTextCase, type FieldTextCase } from './field-instruction';
import type { FormFieldChoice, FormFieldInfo } from './form-fill-fields';

/** What a host can give a field: text, a check box state, or a number. */
export type FormFieldValue = string | boolean | number;

/** A value made ready for a field, or the reason it cannot go in. */
export type PreparedValue =
	| { text: string; date?: Date }
	| { checked: boolean }
	| { choice: number }
	| { clear: true }
	| { reason: string };

/** A calendar day given as `YYYY-MM-DD` (a time after it is ignored). */
export function parseDay(value: string): Date | undefined {
	const match = /^(\d{4})-(\d{2})-(\d{2})(?:T[\d:.]+(?:Z|[+-]\d{2}:\d{2})?)?$/u.exec(value.trim());
	if (!match) return undefined;
	const [year, month, day] = [Number(match[1]), Number(match[2]), Number(match[3])];
	const date = new Date(year, month - 1, day);
	return date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day
		? date
		: undefined;
}

const group = (digits: string) => digits.replace(/\B(?=(\d{3})+(?!\d))/gu, ',');

/**
 * A number written with a Word number picture (`0`, `0.00`, `#,##0.00`, `0%`, with text around
 * it, and a second `;` section for negative numbers). Undefined for a picture it cannot read.
 */
export function formatNumber(value: number, picture: string): string | undefined {
	const [positive, negative] = picture.split(';');
	const own = value < 0 && negative !== undefined;
	const match = /^([^#0,.%]*)([#0,]*)(?:\.([#0]+))?(%?)([^#0,.%]*)$/u.exec(
		own ? negative : positive!,
	);
	if (!match || !match[2]) return undefined;
	const [, prefix, whole, fraction = '', percent, suffix] = match as unknown as string[];
	const amount = (own ? -value : value) * (percent ? 100 : 1);
	const required = (fraction!.match(/0/gu) ?? []).length;
	let [integer, decimals = ''] = Math.abs(amount).toFixed(fraction!.length).split('.');
	while (decimals.length > required && decimals.endsWith('0')) decimals = decimals.slice(0, -1);
	if (whole!.includes(',')) integer = group(integer!);
	const sign = amount < 0 ? '-' : '';
	return `${prefix}${sign}${integer}${decimals ? `.${decimals}` : ''}${percent}${suffix}`;
}

/** Word's text formats for a legacy text field (`w:textInput/w:format`). */
const TEXT_CASES: Record<string, FieldTextCase> = {
	UPPERCASE: 'upper',
	LOWERCASE: 'lower',
	'FIRST CAPITAL': 'firstcap',
	'TITLE CASE': 'caps',
};

/** The entry a value names: by its text or value, then the same ignoring case when only one does. */
export function choiceIndex(
	choices: readonly FormFieldChoice[],
	value: string,
): number | undefined {
	const exact = choices.findIndex((item) => item.display === value || item.value === value);
	if (exact >= 0) return exact;
	const folded = value.toLocaleLowerCase();
	const near = choices
		.map((item, index) => ({ item, index }))
		.filter(
			({ item }) =>
				item.display.toLocaleLowerCase() === folded || item.value.toLocaleLowerCase() === folded,
		);
	return near.length === 1 ? near[0]!.index : undefined;
}

const entriesOf = (choices: readonly FormFieldChoice[]) =>
	choices.map((item) => item.display).join(', ');

/** A legacy text field takes no more characters than its maximum length, whatever its type. */
function fits(field: FormFieldInfo, prepared: PreparedValue): PreparedValue {
	if (!('text' in prepared) || !field.maxLength || [...prepared.text].length <= field.maxLength)
		return prepared;
	return { reason: `longer than the field's ${field.maxLength} characters` };
}

/** Checks `value` for `field` and turns it into what the field will hold. */
export const prepareValue = (field: FormFieldInfo, value: FormFieldValue): PreparedValue =>
	fits(field, prepare(field, value));

function prepare(field: FormFieldInfo, value: FormFieldValue): PreparedValue {
	if (field.type === 'checkbox')
		return typeof value === 'boolean' ? { checked: value } : { reason: 'expected true or false' };
	if (typeof value === 'boolean') return { reason: 'expected text, not true or false' };
	const text = typeof value === 'number' ? String(value) : value;
	if (text === '') {
		if (field.source === 'legacy' && field.type === 'dropdown')
			return { reason: 'a drop-down form field always shows one of its entries' };
		return { clear: true };
	}
	if (field.type === 'dropdown') {
		const index = choiceIndex(field.choices ?? [], text);
		if (index !== undefined) return { choice: index };
		return field.allowsOtherText
			? { text }
			: { reason: `not one of the entries: ${entriesOf(field.choices ?? [])}` };
	}
	if (field.type === 'date') {
		const date = parseDay(text);
		if (!date) return { reason: 'expected a date as YYYY-MM-DD' };
		// A control writes its own date text; a legacy field shows the date with its picture.
		const shown =
			field.source === 'control'
				? ''
				: formatWordDate(date, field.format ?? 'M/d/yyyy', field.locale ?? 'en-US');
		return { text: shown, date };
	}
	if (field.type === 'number') {
		if (typeof value !== 'number' && !/^[-+]?(\d+\.?\d*|\.\d+)$/u.test(text.trim()))
			return { reason: 'expected a number' };
		const amount = Number(text);
		return { text: (field.format && formatNumber(amount, field.format)) || String(amount) };
	}
	if (/[\r\n]/u.test(text) && !field.multiLine) return { reason: 'the field takes one line' };
	const textCase = field.source === 'legacy' && field.format ? TEXT_CASES[field.format] : undefined;
	return { text: textCase ? applyFieldTextCase(text, textCase) : text };
}
