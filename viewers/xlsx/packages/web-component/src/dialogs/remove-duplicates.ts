// Remove Duplicates: choose the key columns and whether the first row is a header.
import { columnLabel, getCell } from '@christophervr/xlsx-core';
import { guessHeader, regionOf, target } from '../commands/util.js';
import type { EditorContext } from '../context.js';
import { checkbox, el, fieldset, row } from './fields.js';
import { button, showDialog } from './frame.js';

export function openRemoveDuplicates(ctx: EditorContext): Promise<number | undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const region = regionOf(t);
	const headers = checkbox(ctx, 'My data has headers', guessHeader(t.ws, region));
	const list = el(ctx, 'div');
	const cols: { col: number; input: HTMLInputElement; wrapper: HTMLElement }[] = [];
	const labels = (): void => {
		list.replaceChildren();
		cols.length = 0;
		for (let col = region.start.col; col <= region.end.col; col++) {
			const head = getCell(t.ws, region.start.row, col)?.value;
			const name =
				headers.input.checked && head !== null && head !== undefined && head !== ''
					? String(typeof head === 'object' ? head.error : head)
					: ctx.t('Column {name}', { name: columnLabel(col) });
			const box = el(ctx, 'label', 'xve-check');
			const input = el(ctx, 'input');
			input.type = 'checkbox';
			input.checked = true;
			input.setAttribute('aria-label', name);
			const span = el(ctx, 'span');
			span.textContent = name;
			box.append(input, span);
			list.append(box);
			cols.push({ col, input, wrapper: box });
		}
	};
	labels();
	headers.input.addEventListener('change', labels);
	const all = button(ctx, 'Select All');
	const none = button(ctx, 'Unselect All');
	all.addEventListener('click', () => cols.forEach((c) => (c.input.checked = true)));
	none.addEventListener('click', () => cols.forEach((c) => (c.input.checked = false)));
	return showDialog<number>(ctx, {
		name: 'remove-duplicates',
		heading: 'Remove Duplicates',
		body: [row(ctx, all, none), headers.wrapper, fieldset(ctx, 'Columns', list)],
		opened: () => all.focus(),
		submit: () => {
			const chosen = cols.filter((c) => c.input.checked).map((c) => c.col);
			if (!chosen.length) {
				ctx.toast(ctx.t('Select at least one column.'), 'warning');
				return undefined;
			}
			const hasHeader = headers.input.checked;
			const { removed, remaining } = t.session.removeDuplicates(t.sheet, region, chosen, hasHeader);
			ctx.toast(
				ctx.t('{removed} duplicate values found and removed; {kept} unique values remain.', {
					removed,
					kept: remaining,
				}),
				'info',
			);
			return removed;
		},
	});
}
