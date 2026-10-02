// Ctrl+; and Ctrl+Shift+; : the current date or time as a constant. In Ready mode the active cell
// gets the value (with a date or time format when it has none); while editing, the text goes in
// at the caret, as in Excel.
import { dateToSerial, formatValue, isDateFormat } from '@christophervr/xlsx-core';
import type { Command } from '../commands.js';
import type { EditorContext } from '../context.js';
import { activeStyle, editing, target } from './util.js';

const FORMATS = { date: 'm/d/yyyy', time: 'h:mm AM/PM' } as const;

/** Today's date or the current time (to the minute) as a serial, in local wall time. */
export function nowSerial(kind: 'date' | 'time', now = new Date()): number {
	const wall = new Date(
		Date.UTC(now.getFullYear(), now.getMonth(), now.getDate(), now.getHours(), now.getMinutes()),
	);
	const serial = dateToSerial(wall);
	return kind === 'date' ? Math.floor(serial) : serial - Math.floor(serial);
}

function insertNow(ctx: EditorContext, kind: 'date' | 'time'): void {
	const t = target(ctx);
	if (!t) return;
	const serial = nowSerial(kind);
	const field = ctx.root.activeElement;
	if (ctx.grid()?.isEditing() && field instanceof HTMLTextAreaElement) {
		const text = formatValue(serial, FORMATS[kind], { date1904: t.workbook.date1904 }).text;
		field.setRangeText(text, field.selectionStart, field.selectionEnd, 'end');
		field.dispatchEvent(new Event('input', { bubbles: true }));
		return;
	}
	const { row, col } = t.active;
	const current = activeStyle(ctx)?.numFmt ?? 'General';
	t.session.batch(kind === 'date' ? 'Insert date' : 'Insert time', () => {
		t.session.setCellValue(t.sheet, row, col, serial);
		if (!isDateFormat(current))
			t.session.applyStyle(t.sheet, [{ start: t.active, end: t.active }], {
				numFmt: FORMATS[kind],
			});
	});
}

export function insertNowCommands(): Command[] {
	return (['date', 'time'] as const).map((kind) =>
		editing({
			id: `edit.insert-${kind}`,
			label: kind === 'date' ? 'Insert the current date' : 'Insert the current time',
			shortcut: kind === 'date' ? 'Ctrl+;' : 'Ctrl+Shift+;',
			run: (ctx) => insertNow(ctx, kind),
		}),
	);
}
