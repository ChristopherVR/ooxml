import { fail } from './package-common';
import { analyzeVisioFormula, parseVisioFormula } from './formula';

/**
 * Insert a text field (`<fld>` plus a Field section row) into a local shape's text. The Value
 * formula is evaluated now for the cached text; offsets count each existing field as its cached
 * text, as the shape's source text does.
 */
export interface VisioTextFieldInsertEdit {
	type: 'insert-text-field';
	pageId: string;
	shapeId: string;
	/** Field Value formula: NOW(), DOCLASTSAVE(), TITLE(), CREATOR(), PAGENAME(), PAGENUMBER(),
	 * PAGECOUNT(), Width, Height, Angle or a numeric custom formula on this shape. */
	formula: string;
	/** Format picture from VISIO_FIELD_FORMATS (or the same subset); empty is general. */
	format?: string;
	/** Zero-based insertion offset in the source text; omitted appends. */
	offset?: number;
	/** Source text the offset refers to; required with `offset` so stale offsets refuse. */
	expectedText?: string;
}

/** Atomically update one live local Field row and its unique <fld>, retaining its IX. */
export interface VisioTextFieldUpdateEdit {
	type: 'update-text-field';
	pageId: string;
	shapeId: string;
	fieldIndex: string;
	formula: string;
	/** Omitted or empty format means General, as for insertion. */
	format?: string;
	/** Exact source text, cached field text and Field row key from the parsed scene. */
	expectedText: string;
	expectedCached: string;
	expectedSourceKey: string;
}

export function snapshotTextFieldUpdate(edit: VisioTextFieldUpdateEdit): VisioTextFieldUpdateEdit {
	const insert = snapshotTextFieldInsert({
		type: 'insert-text-field',
		pageId: edit.pageId,
		shapeId: edit.shapeId,
		formula: edit.formula,
		...(edit.format === undefined ? {} : { format: edit.format }),
	});
	if (
		analyzeVisioFormula(parseVisioFormula(insert.formula)).references.some((ref) =>
			/^fields?\./i.test(ref.cell),
		)
	)
		fail('UNSUPPORTED_TEXT_FIELD', 'Existing field formulas cannot reference Field rows.');
	if (
		typeof edit.fieldIndex !== 'string' ||
		!/^(0|[1-9]\d{0,9})$/.test(edit.fieldIndex) ||
		Number(edit.fieldIndex) > 0xffffffff
	)
		fail('INVALID_EDIT', 'An existing field needs a canonical row index.');
	if (
		[edit.expectedText, edit.expectedCached, edit.expectedSourceKey].some(
			(value) => typeof value !== 'string' || value.length > 1_000_000,
		) ||
		!edit.expectedSourceKey
	)
		fail('INVALID_EDIT', 'An existing field needs its source text and row identity.');
	return {
		type: 'update-text-field',
		pageId: insert.pageId,
		shapeId: insert.shapeId,
		fieldIndex: edit.fieldIndex,
		formula: insert.formula,
		format: insert.format!,
		expectedText: edit.expectedText,
		expectedCached: edit.expectedCached,
		expectedSourceKey: edit.expectedSourceKey,
	};
}

export function snapshotTextFieldInsert(edit: VisioTextFieldInsertEdit): VisioTextFieldInsertEdit {
	const { pageId, shapeId } = edit;
	if ([pageId, shapeId].some((id) => typeof id !== 'string' || !id || id.length > 256))
		fail('INVALID_EDIT', 'Invalid text field target.');
	const formula = edit.formula;
	if (typeof formula !== 'string' || !formula.trim() || formula.length > 1024)
		fail('INVALID_EDIT', 'A text field needs a formula of at most 1024 characters.');
	let dynamic = true;
	try {
		dynamic = analyzeVisioFormula(parseVisioFormula(formula)).dynamic;
	} catch {
		fail('INVALID_EDIT', 'The text field formula is not valid ShapeSheet syntax.');
	}
	if (dynamic)
		fail('INVALID_EDIT', 'Text field formulas cannot use dynamic references (INDIRECT, REF...).');
	const format = edit.format ?? '';
	if (typeof format !== 'string' || format.length > 256 || /[\u0000-\u001f\u007f]/.test(format))
		fail('INVALID_EDIT', 'Invalid text field format.');
	const result: VisioTextFieldInsertEdit = {
		type: 'insert-text-field',
		pageId,
		shapeId,
		formula,
		format,
	};
	if (edit.offset !== undefined) {
		if (
			!Number.isSafeInteger(edit.offset) ||
			edit.offset < 0 ||
			typeof edit.expectedText !== 'string' ||
			edit.offset > edit.expectedText.length ||
			edit.expectedText.length > 1_000_000
		)
			fail('INVALID_EDIT', 'A field offset needs the source text it refers to.');
		result.offset = edit.offset;
		result.expectedText = edit.expectedText;
	}
	return result;
}
