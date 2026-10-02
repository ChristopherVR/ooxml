// Show Comments: the sheet's comments (cell, author, text); activating one selects its cell.
import { formatAddress } from '@christophervr/xlsx-core';
import type { EditorContext } from '../context.js';
import { listBox, text } from './fields.js';
import { showDialog } from './frame.js';
import { selectOn } from './go-to.js';

export function openCommentsList(ctx: EditorContext): Promise<string | undefined> {
	const ws = ctx.workbook()?.sheets[ctx.activeSheet()];
	const comments = [...(ws?.comments ?? [])].sort(
		(a, b) => a.address.row - b.address.row || a.address.col - b.address.col,
	);
	let close: (() => void) | undefined;
	let chosen: string | undefined;
	const activate = (value: string): void => {
		const comment = comments[Number(value)];
		if (!comment) return;
		const at = { ...comment.address };
		chosen = formatAddress(at);
		selectOn(ctx, ctx.activeSheet(), { start: at, end: at });
		close?.();
	};
	const list = listBox(ctx, 'Comments', undefined, activate);
	list.setItems(
		comments.map(
			(c, i) =>
				[
					String(i),
					`${formatAddress(c.address)}  ${c.author}: ${c.text.replace(/\s+/g, ' ')}`,
				] as const,
		),
	);
	for (const option of list.element.querySelectorAll<HTMLElement>('[role="option"]'))
		option.addEventListener(
			'click',
			() => option.dataset.value !== undefined && activate(option.dataset.value),
		);
	const first = comments.length ? '0' : undefined;
	if (first) list.select(first);
	const body = comments.length
		? [list.element]
		: [text(ctx, 'There are no comments on this sheet.')];
	return showDialog<string>(
		ctx,
		{
			name: 'comments-list',
			heading: 'Comments',
			okLabel: null,
			body,
			opened: () => list.element.focus(),
		},
		(open) => {
			close = open.close;
		},
	).then((result) => result ?? chosen);
}
