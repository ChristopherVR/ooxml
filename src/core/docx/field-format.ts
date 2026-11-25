// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// A field instruction read into its arguments, its own switches and the general formatting
// switches (ECMA-376 Part 1, 17.16.4.1): `\*` text case and number formats, `\@` date pictures.
// What these results look like was checked against Word's own field results.
import {
	applyFieldTextCase,
	fieldTokens,
	isFieldTextCase,
	type FieldTextCase,
} from './field-instruction';
import { formatWordDate } from './field-date';
import { formatListNumber } from './numbering-format';

/** One `\*` format of a field, applied in the order the instruction gives them. */
export type FieldFormatStep =
	| { kind: 'case'; textCase: FieldTextCase }
	/** A `w:numFmt` name `formatListNumber` draws, or `arabicDash` (`- 42 -`). */
	| { kind: 'number'; numFmt: string };

/** A field's general formatting switches. */
export interface FieldFormat {
	steps: FieldFormatStep[];
	/** `\* CHARFORMAT`: the result takes the formatting of the field name's first letter. */
	charFormat?: true;
	/** `\@` date-time picture. */
	datePicture?: string;
	/** `\#` numeric picture, for callers that evaluate that switch. */
	numberPicture?: string;
}

/** A field instruction: its name, arguments, own switches and general formatting. */
export interface ParsedField {
	/** The field type, upper-cased (`DOCPROPERTY`). */
	name: string;
	args: string[];
	/** Own switches by letter: their argument, or `true` for a switch that takes none. */
	switches: Map<string, string | true>;
	format: FieldFormat;
}

/**
 * The `\*` number formats Word draws: a lower-case first letter gives lower-case numerals and
 * letters (`roman` xlii, `Roman` and `ROMAN` XLII), as Word's own results show.
 */
function numberFormat(argument: string): string | undefined {
	const lower = argument.toLowerCase();
	const upper = argument[0] !== argument[0]?.toLowerCase();
	if (lower === 'arabic') return 'decimal';
	if (lower === 'roman') return upper ? 'upperRoman' : 'lowerRoman';
	if (lower === 'alphabetic') return upper ? 'upperLetter' : 'lowerLetter';
	if (lower === 'cardtext') return 'cardinalText';
	if (lower === 'ordtext') return 'ordinalText';
	if (lower === 'ordinal') return 'ordinal';
	if (lower === 'hex') return 'hex';
	if (lower === 'arabicdash') return 'arabicDash';
	return undefined;
}

/**
 * Reads an instruction. `own` lists the field's switches: `'arg'` for one that takes an
 * argument, `'flag'` for one that does not. Undefined for an empty instruction, an unknown switch
 * or a general switch that is not evaluated here (`\#` number pictures, a `\*` format Word does
 * not know): such a field keeps the result it was saved with.
 */
export function parseFieldInstruction(
	instr: string,
	own: Readonly<Record<string, 'arg' | 'flag'>> = {},
	options: { numberPictures?: boolean } = {},
): ParsedField | undefined {
	const tokens = fieldTokens(instr);
	const head = tokens[0];
	if (head?.kind !== 'text') return undefined;
	const field: ParsedField = {
		name: head.value.toUpperCase(),
		args: [],
		switches: new Map(),
		format: { steps: [] },
	};
	for (let index = 1; index < tokens.length; index++) {
		const token = tokens[index]!;
		if (token.kind === 'text') {
			field.args.push(token.value);
			continue;
		}
		const argument = () => {
			const next = tokens[index + 1];
			if (next?.kind !== 'text') return undefined;
			index++;
			return next.value;
		};
		if (token.value === '*') {
			const format = argument();
			if (format === undefined) return undefined;
			const lower = format.toLowerCase();
			if (lower === 'mergeformat') continue;
			if (lower === 'charformat') field.format.charFormat = true;
			else if (isFieldTextCase(lower)) field.format.steps.push({ kind: 'case', textCase: lower });
			else {
				const numFmt = numberFormat(format);
				if (!numFmt) return undefined;
				field.format.steps.push({ kind: 'number', numFmt });
			}
		} else if (token.value === '@') {
			const picture = argument();
			if (picture === undefined) return undefined;
			field.format.datePicture = picture;
		} else if (token.value === '#' && options.numberPictures) {
			const picture = argument();
			if (picture === undefined) return undefined;
			field.format.numberPicture = picture;
		} else if (token.value === '!') continue;
		else if (own[token.value] === 'arg') {
			const value = argument();
			field.switches.set(token.value, value ?? true);
		} else if (own[token.value] === 'flag') field.switches.set(token.value, true);
		else return undefined;
	}
	return field;
}

/** `value` in a `\*` number format; Word rounds a fraction to the nearest whole number. */
export function formatFieldNumber(value: number, numFmt: string): string {
	const whole = Math.round(value);
	if (numFmt === 'arabicDash') return `- ${whole} -`;
	return formatListNumber(numFmt, whole);
}

/**
 * A field result with its `\*` formats applied in order: a number format redraws a result that
 * is a number (other text is left as it is), a text case changes its letters.
 */
export function applyFieldFormat(text: string, format: FieldFormat): string {
	let result = text;
	for (const step of format.steps) {
		if (step.kind === 'case') result = applyFieldTextCase(result, step.textCase);
		else if (/^\s*-?\d+(\.\d+)?\s*$/.test(result))
			result = formatFieldNumber(Number(result), step.numFmt);
	}
	return result;
}

/** A date result: with a `\@` picture, `date` drawn with it; undefined without one. */
export function fieldDateResult(date: Date, format: FieldFormat): string | undefined {
	return format.datePicture === undefined
		? undefined
		: applyFieldFormat(formatWordDate(date, format.datePicture), format);
}
