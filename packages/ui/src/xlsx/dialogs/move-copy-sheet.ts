// Move or Copy: reorder the active sheet within this workbook, optionally as a copy.
import type { EditorContext } from '../context.js';
import { checkbox, field, listBox, select, text } from './fields.js';
import { showDialog } from './frame.js';

export const END = 'end';

/** The index a sheet ends up at when moved before `before` (sheet count = move to end). */
export function moveTarget(from: number, before: number): number {
	return before > from ? before - 1 : before;
}

export function openMoveCopySheet(ctx: EditorContext): Promise<number | undefined> {
	const session = ctx.session();
	if (!session) return Promise.resolve(undefined);
	if (session.workbook.structureLocked) {
		ctx.toast(ctx.t('Workbook structure is protected.'), 'warning');
		return Promise.resolve(undefined);
	}
	const from = ctx.activeSheet();
	const book = select(ctx, [['this', '(current workbook)']], 'this');
	book.disabled = true;
	const list = listBox(ctx, 'Before sheet:');
	list.setItems([
		...session.workbook.sheets.map((s, i) => [String(i), s.name] as const),
		[END, ctx.t('(move to end)')] as const,
	]);
	list.select(String(Math.min(from + 1, session.workbook.sheets.length - 1)));
	const copy = checkbox(ctx, 'Create a copy');
	return showDialog<number>(ctx, {
		name: 'move-copy-sheet',
		heading: 'Move or Copy',
		body: [
			text(ctx, 'Move selected sheets', 'xve-field-label'),
			field(ctx, 'To book:', book),
			text(ctx, 'Before sheet:', 'xve-field-label'),
			list.element,
			copy.wrapper,
		],
		opened: () => list.element.focus(),
		submit: () => {
			const choice = list.value();
			if (choice === undefined) return undefined;
			const count = session.workbook.sheets.length;
			const before = choice === END ? count : Number(choice);
			let result = from;
			session.batch(copy.input.checked ? 'Copy sheet' : 'Move sheet', () => {
				if (copy.input.checked) {
					const dup = session.duplicateSheet(from);
					// Indices after the original moved up by one when the copy was inserted.
					const beforeNow = before > from ? before + 1 : before;
					const to = Math.min(moveTarget(dup, beforeNow), session.workbook.sheets.length - 1);
					if (to !== dup) session.moveSheet(dup, to);
					result = to;
				} else {
					const to = Math.min(moveTarget(from, before), count - 1);
					if (to !== from) session.moveSheet(from, to);
					result = to;
				}
			});
			ctx.setActiveSheet(result);
			return result;
		},
	});
}
