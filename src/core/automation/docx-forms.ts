import type { DocumentProtection } from '../docx/form-field-model';
import { listFormFields, type FormFieldInfo } from '../docx/form-fill-fields';
import type { FormFieldValue } from '../docx/form-fill-values';
import { fillFormFields, type FormFillOptions, type FormFillResult } from '../docx/form-fill-write';
import { loadDocx } from '../docx/parse';

/** Headless form filling for Word documents, shared by MCP and other automation adapters. */
export async function listDocxFormFields(bytes: Uint8Array): Promise<{
	fields: FormFieldInfo[];
	/** `w:documentProtection`, when the document restricts editing (for example to forms). */
	protection?: DocumentProtection;
	warnings: string[];
}> {
	const { model } = await loadDocx(bytes);
	return {
		fields: listFormFields(model),
		...(model.protection && { protection: model.protection }),
		warnings: model.warnings,
	};
}

export interface DocxFormFillOptions extends FormFillOptions {
	/** Write the values that fit even when others are refused (by default nothing is written). */
	partial?: boolean;
}

/**
 * Sets form fields by the ids `listDocxFormFields` gives and returns the saved document. Values
 * are checked against their fields; when one is refused nothing is written (unless `partial`),
 * and the error lists every refusal. Parts the values do not reach are kept byte for byte.
 */
export async function fillDocxFormFields(
	bytes: Uint8Array,
	values: Readonly<Record<string, FormFieldValue>>,
	options: DocxFormFillOptions = {},
): Promise<{ bytes: Uint8Array; results: FormFillResult[]; warnings: string[] }> {
	const loaded = await loadDocx(bytes);
	const { model, results } = fillFormFields(loaded.model, values, options);
	const refused = results.filter((result) => !result.ok);
	if (refused.length && !options.partial)
		throw new Error(
			`No field was filled: ${refused.map((result) => `${result.id}: ${result.reason}`).join('; ')}`,
		);
	const warnings = [...model.warnings];
	if (model.trackChanges)
		warnings.push('Track Changes is on in this document; the filled values are not tracked.');
	return { bytes: await loaded.save(model), results, warnings };
}
