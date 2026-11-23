// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// The fields whose results this package computes from the document itself: DOCPROPERTY, the
// AUTONUM fields, LISTNUM and SYMBOL. Layout draws these results, Update Fields writes them into the editor and a
// save writes them back as the fields' cached results, so all three read them from here.
import {
	AUTO_NUMBER_FIELDS,
	autoNumberResult,
	autoNumberStates,
	type AutoNumberState,
} from './auto-number-fields';
import {
	DOC_INFO_FIELD_MAP,
	NO_PROPERTY_NAME,
	UNKNOWN_PROPERTY,
	docInfoFieldResult,
	docPropertyResult,
	fieldProperties,
} from './doc-property-field';
import { ifFieldResult } from './conditional-field';
import { ZERO_DIVIDE_ERROR, formulaFieldResult } from './formula-field';
import type { ListNumberFields } from './list-number-fields';
import { computeListNumbering } from './numbering-format';
import { symbolFieldResult } from './symbol-field';
import { fieldName } from './field-runs';
import { withFieldResults, type FieldResultSpan } from './field-result-spans';
import { withStoryBlocks } from './field-stories';
import type { DocumentPropertySet } from '../opc/properties/types';
import type { DocumentModel, Paragraph } from './model';

/** What a computed field shows, and the font it draws in when that is not the run's own. */
export interface ComputedFieldResult {
	text: string;
	/** `\* CHARFORMAT`: formatted like the first letter of the field name. */
	charFormat?: boolean;
	fontFamily?: string;
	fontSizePt?: number;
}

/** The computed results of one document's fields. */
export interface ComputedFields {
	/**
	 * What `field` of `paragraph` shows; undefined for a field this package does not compute, or
	 * one whose switches it does not evaluate: such a field keeps its saved result.
	 */
	result(
		field: Pick<FieldResultSpan, 'from' | 'instr'>,
		paragraph: Paragraph,
	): ComputedFieldResult | undefined;
}

const CHARFORMAT = /\*\s*charformat\b/i;

const DOC_INFO_NAMES = Object.keys(DOC_INFO_FIELD_MAP);

const normalizedFieldName = (instr: string): string => {
	const name = fieldName(instr);
	return name.startsWith('=') ? '=' : name;
};

/** The names of the fields `computedFields` evaluates. */
export const COMPUTED_FIELD_NAMES: ReadonlySet<string> = new Set([
	'DOCPROPERTY',
	...DOC_INFO_NAMES,
	'IF',
	'=',
	...AUTO_NUMBER_FIELDS,
	'LISTNUM',
	'SYMBOL',
]);

/**
 * Reads `model`'s computed fields; what each needs is worked out once, when first asked for. A
 * DOCPROPERTY that names no property keeps its result unless `errors` asks for Word's error text,
 * as an explicit Update Fields shows it: the layout and a save never turn a result into an error.
 */
export function computedFields(
	model: DocumentModel,
	options: { errors?: boolean } = {},
): ComputedFields {
	let properties: DocumentPropertySet | undefined;
	let autoNumbers: Map<string, AutoNumberState> | undefined;
	let listNumbers: ListNumberFields | undefined;
	return {
		result(field, paragraph) {
			const name = normalizedFieldName(field.instr);
			if (name === 'DOCPROPERTY') {
				const text = docPropertyResult(field.instr, (properties ??= fieldProperties(model)));
				const error = text === UNKNOWN_PROPERTY || text === NO_PROPERTY_NAME;
				return text === undefined || (error && !options.errors) ? undefined : { text };
			}
			if (name in DOC_INFO_FIELD_MAP) {
				const text = docInfoFieldResult(field.instr, (properties ??= fieldProperties(model)));
				return text === undefined ? undefined : { text };
			}
			if (name === 'IF') {
				const text = ifFieldResult(field.instr);
				return text === undefined ? undefined : { text };
			}
			if (name === '=') {
				const text = formulaFieldResult(field.instr);
				if (text === undefined || (text === ZERO_DIVIDE_ERROR && !options.errors))
					return undefined;
				return { text };
			}
			if (AUTO_NUMBER_FIELDS.has(name)) {
				const state = (autoNumbers ??= autoNumberStates(model)).get(paragraph.id);
				const text = state && autoNumberResult(field.instr, state);
				return text === undefined ? undefined : { text };
			}
			if (name === 'SYMBOL') {
				const symbol = symbolFieldResult(field.instr);
				return (
					symbol && {
						text: symbol.text,
						...(symbol.fontFamily ? { fontFamily: symbol.fontFamily } : {}),
						...(symbol.fontSizePt ? { fontSizePt: symbol.fontSizePt } : {}),
					}
				);
			}
			if (name === 'LISTNUM') {
				listNumbers ??= computeListNumbering(model).listNumbers;
				const text = listNumbers.result(paragraph.id, field.from);
				return text === undefined ? undefined : { text };
			}
			return undefined;
		},
	};
}

/** The computed fields whose results do not depend on where they are in the main document. */
const STORY_FIELD_NAMES: ReadonlySet<string> = new Set([
	'DOCPROPERTY',
	...DOC_INFO_NAMES,
	'IF',
	'=',
	'SYMBOL',
]);

/**
 * `model` with the cached result of every unlocked computed field brought up to date, in the
 * headers and footers and, with `body`, the main document. A field saved without a result
 * (Word saves AUTONUM, LISTNUM and SYMBOL so, and draws what they show) is left without one.
 * The counted fields (AUTONUM, LISTNUM) count the main document only, so a header or footer
 * keeps theirs. Returns `model` itself when nothing changed.
 */
export function refreshComputedFields(
	model: DocumentModel,
	options: { body?: boolean } = {},
): DocumentModel {
	const fields = computedFields(model);
	const story = (names: ReadonlySet<string>) => (blocks: Parameters<typeof withFieldResults>[0]) =>
		withFieldResults(blocks, (field, paragraph) => {
			if (!names.has(normalizedFieldName(field.instr))) return undefined;
			const result = fields.result(field, paragraph);
			return result && { text: result.text, charFormat: CHARFORMAT.test(field.instr) };
		});
	return withStoryBlocks(
		model,
		story(STORY_FIELD_NAMES),
		options.body ? story(COMPUTED_FIELD_NAMES) : undefined,
	);
}
