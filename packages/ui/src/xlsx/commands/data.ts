// Data tab: Data Tools (Text to Columns, Remove Duplicates, Data Validation) and Outline (group
// and ungroup rows or columns). Sort & Filter commands live in sort-filter.ts.
import type { Command } from '../commands.js';
import { icon } from './icons.js';
import { editing, rowsOf, target, wholeColumns, wholeRows } from './util.js';

export function dataCommands(): Command[] {
	return [
		editing({
			id: 'data.text-to-columns',
			label: 'Text to Columns',
			icon: icon('textColumns'),
			run: (ctx) => void ctx.dialogs.open('text-to-columns'),
		}),
		editing({
			id: 'data.remove-duplicates',
			label: 'Remove Duplicates',
			icon: icon('removeDuplicates'),
			run: (ctx) => void ctx.dialogs.open('remove-duplicates'),
		}),
		editing({
			id: 'data.validation',
			label: 'Data Validation...',
			icon: icon('validation'),
			run: (ctx) => void ctx.dialogs.open('data-validation'),
		}),
		...(['group', 'ungroup'] as const).map((mode) =>
			editing({
				id: `data.${mode}`,
				label: mode === 'group' ? 'Group' : 'Ungroup',
				icon: icon(mode),
				shortcut: mode === 'group' ? 'Alt+Shift+Right' : 'Alt+Shift+Left',
				enabled: (ctx) => {
					const t = target(ctx);
					if (!t) return false;
					if (mode === 'group') return true;
					const r = t.range;
					return (
						rowsOf(t.ranges).some((row) => (t.ws.rowInfo.get(row)?.outlineLevel ?? 0) > 0) ||
						t.ws.columns.some(
							(c) => (c.outlineLevel ?? 0) > 0 && c.max >= r.start.col && c.min <= r.end.col,
						)
					);
				},
				run: async (ctx, arg) => {
					const t = target(ctx);
					if (!t) return;
					const r = t.range;
					const axis =
						arg === 'rows' || arg === 'cols'
							? arg
							: wholeColumns(r)
								? 'cols'
								: wholeRows(r)
									? 'rows'
									: await ctx.dialogs.open<'rows' | 'cols'>(mode);
					if (!axis) return;
					const { session, sheet } = t;
					if (axis === 'rows') {
						if (mode === 'group') session.groupRows(sheet, r.start.row, r.end.row);
						else session.ungroupRows(sheet, r.start.row, r.end.row);
					} else if (mode === 'group') session.groupColumns(sheet, r.start.col, r.end.col);
					else session.ungroupColumns(sheet, r.start.col, r.end.col);
				},
			}),
		),
	];
}
