// Paste Special: what to paste (all, formulas, values, formats) and Transpose. Operations, Skip
// blanks, All except borders and Column widths are shown disabled: the core paste has no such modes.
import type { PasteMode } from '@christophervr/xlsx-core';
import { pasteWith } from '../commands/clipboard.js';
import type { EditorContext } from '../context.js';
import { checkbox, radios, row, text } from './fields.js';
import { showDialog } from './frame.js';

const PASTE: ReadonlyArray<readonly [string, string]> = [
	['all', 'All'],
	['formulas', 'Formulas'],
	['values', 'Values'],
	['formats', 'Formats'],
	['noBorders', 'All except borders'],
	['widths', 'Column widths'],
];
const OPERATIONS: ReadonlyArray<readonly [string, string]> = [
	['none', 'None'],
	['add', 'Add'],
	['subtract', 'Subtract'],
	['multiply', 'Multiply'],
	['divide', 'Divide'],
];

export function openPasteSpecial(ctx: EditorContext): Promise<PasteMode | undefined> {
	const paste = radios(ctx, 'Paste', PASTE, 'all');
	for (const input of paste.inputs)
		if (input.value === 'noBorders' || input.value === 'widths') input.disabled = true;
	const operation = radios(ctx, 'Operation', OPERATIONS, 'none');
	for (const input of operation.inputs) if (input.value !== 'none') input.disabled = true;
	const skip = checkbox(ctx, 'Skip blanks');
	skip.input.disabled = true;
	const transpose = checkbox(ctx, 'Transpose');
	return showDialog<PasteMode>(ctx, {
		name: 'paste-special',
		heading: 'Paste Special',
		body: [
			row(ctx, paste.element, operation.element),
			row(ctx, skip.wrapper, transpose.wrapper),
			text(
				ctx,
				'Operations, skip blanks and transpose combined with values are not supported yet.',
			),
		],
		opened: () => paste.inputs[0]?.focus(),
		submit: () => {
			const mode: PasteMode = transpose.input.checked ? 'transpose' : (paste.get() as PasteMode);
			pasteWith(ctx, mode);
			return mode;
		},
	});
}
