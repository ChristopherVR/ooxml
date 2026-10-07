// Paste Special gathers options; the requesting command applies them once.
import type { PasteOptions } from 'ooxml-core/xlsx';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { checkbox, radios, row } from './fields.js';
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

export function openPasteSpecial(ctx: EditorContext): Promise<PasteOptions | undefined> {
	const paste = radios(ctx, 'Paste', PASTE, 'all');
	const operation = radios(ctx, 'Operation', OPERATIONS, 'none');
	const skip = checkbox(ctx, 'Skip blanks');
	const transpose = checkbox(ctx, 'Transpose');
	return showDialog<PasteOptions>(ctx, {
		name: 'paste-special',
		heading: 'Paste Special',
		body: [row(ctx, paste.element, operation.element), row(ctx, skip.wrapper, transpose.wrapper)],
		opened: () => paste.inputs[0]?.focus(),
		submit: () => {
			return {
				mode: paste.get() as NonNullable<PasteOptions['mode']>,
				transpose: transpose.input.checked,
				skipBlanks: skip.input.checked,
				operation: operation.get() as NonNullable<PasteOptions['operation']>,
			};
		},
	});
}
