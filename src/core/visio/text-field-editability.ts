import { attribute, children, readCells } from './sheet';
import { assertFormattingText, formattingRows } from './edit-formatting-rows';
import { textFieldSourceKey } from './text-field-source';
import { fail } from './package-common';

/** Static source restrictions shared by chooser projection and atomic writer. */
export function assertEditableFieldRow(existing: Element): void {
	for (const name of ['Calendar', 'ObjectKind']) {
		const cell = children(existing, 'Cell').find((cell) => attribute(cell, 'N') === name);
		if (cell && (attribute(cell, 'V') !== '0' || cell.hasAttribute('F') || cell.hasAttribute('E')))
			fail(
				'UNSUPPORTED_TEXT_FIELD',
				'Nonstandard field calendar or object profiles cannot be edited.',
			);
	}
	for (const cell of children(existing, 'Cell')) {
		const name = attribute(cell, 'N') ?? '';
		if (
			/^(value|format|type|calendar|objectkind)$/i.test(name) &&
			!['Value', 'Format', 'Type', 'Calendar', 'ObjectKind'].includes(name)
		)
			fail('UNSUPPORTED_TEXT_FIELD', 'Noncanonical field cell names cannot be edited.');
		if (
			['Value', 'Format', 'Type'].includes(attribute(cell, 'N') ?? '') &&
			(cell.hasAttribute('E') || /^\s*(Inh|GUARD\s*\()/i.test(attribute(cell, 'F') ?? ''))
		)
			fail('UNSUPPORTED_TEXT_FIELD', 'Protected or inherited field cells cannot be edited.');
	}
}

export function assertEditableFieldShape(shape: Element): void {
	if (
		shape.localName !== 'Shape' ||
		(shape.parentNode?.parentNode as Element | undefined)?.localName !== 'PageContents' ||
		shape.hasAttribute('Master') ||
		shape.hasAttribute('MasterShape') ||
		attribute(shape, 'Type') === 'Group' ||
		['Shapes', 'ForeignData', 'Rel'].some((name) => children(shape, name).length)
	)
		fail('UNSUPPORTED_TEXT_FIELD', 'Existing field edits require a top-level plain local shape.');
}

/** Conservative source-backed choices; evaluation, dependencies and staleness still check at write. */
export function editableTextFieldIndices(
	text: Element | undefined,
	allowed: boolean,
): ReadonlySet<string> {
	const result = new Set<string>();
	if (!allowed || !text) return result;
	const shape = text.parentNode as Element;

	try {
		assertEditableFieldShape(shape);
		assertFormattingText(shape);
		const rows = formattingRows(shape, 'Field');
		const markers = children(text, 'fld');
		const counts = new Map<string, number>();
		for (const marker of markers) {
			const index = attribute(marker, 'IX')!;
			counts.set(index, (counts.get(index) ?? 0) + 1);
		}
		for (const [index, row] of rows) {
			if (counts.get(index) !== 1 || textFieldSourceKey(readCells(row)) === undefined) continue;
			try {
				assertEditableFieldRow(row);
				result.add(index);
			} catch {
				/* Static unsupported row stays visible but cannot be selected for editing. */
			}
		}
	} catch {
		/* Ambiguous or noncanonical source markup has no editable fields. */
	}
	return result;
}
