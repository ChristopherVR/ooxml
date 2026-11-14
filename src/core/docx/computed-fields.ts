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
import { docPropertyResult, fieldProperties } from './doc-property-field';
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

/** The names of the fields `computedFields` evaluates. */
export const COMPUTED_FIELD_NAMES: ReadonlySet<string> = new Set([
	'DOCPROPERTY',
	...AUTO_NUMBER_FIELDS,
	'LISTNUM',
	'SYMBOL',
]);

/** Reads `model`'s computed fields; what each needs is worked out once, when first asked for. */
export function computedFields(model: DocumentModel): ComputedFields {
	let properties: DocumentPropertySet | undefined;
	let autoNumbers: Map<string, AutoNumberState> | undefined;
	let listNumbers: ListNumberFields | undefined;
	return {
		result(field, paragraph) {
			const name = fieldName(field.instr);
			if (name === 'DOCPROPERTY') {
				const text = docPropertyResult(field.instr, (properties ??= fieldProperties(model)));
				return text === undefined ? undefined : { text };
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

/**
 * `model` with the cached result of every unlocked computed field brought up to date, in the
 * headers and footers and, with `body`, the main document. A field saved without a result
 * (Word saves AUTONUM, LISTNUM and SYMBOL so, and draws what they show) is left without one.
 * Returns `model` itself when nothing changed.
 */
export function refreshComputedFields(
	model: DocumentModel,
	options: { body?: boolean } = {},
): DocumentModel {
	const fields = computedFields(model);
	const story = (blocks: Parameters<typeof withFieldResults>[0]) =>
		withFieldResults(blocks, (field, paragraph) => {
			if (!COMPUTED_FIELD_NAMES.has(fieldName(field.instr))) return undefined;
			const result = fields.result(field, paragraph);
			return result && { text: result.text, charFormat: CHARFORMAT.test(field.instr) };
		});
	return withStoryBlocks(model, story, options.body ? story : undefined);
}
