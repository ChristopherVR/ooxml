// Show Comments: the sheet's comment threads on the shared comments pane (navigation list only);
// choosing one selects its cell.
import { formatAddress, type Comment } from 'ooxml-core/xlsx';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { defineCommentsPane, type OfficeUiCommentsPane } from '../../comments/comments-pane';
import type { OfficeCommentThread } from '../../comments/types';
import { showDialog } from './frame';
import { selectOn } from './go-to';

/** Excel notes and threaded comments as neutral threads, in sheet order, keyed by address. */
export function excelCommentThreads(comments: readonly Comment[]): OfficeCommentThread[] {
	return [...comments]
		.sort((a, b) => a.address.row - b.address.row || a.address.col - b.address.col)
		.map((comment) => {
			const id = formatAddress(comment.address);
			return {
				id,
				anchorLabel: id,
				resolved: false,
				comments: [
					{ id: `${id}:0`, author: comment.author, text: comment.text },
					...(comment.replies ?? []).map((reply, index) => ({
						id: `${id}:${index + 1}`,
						author: reply.author,
						text: reply.text,
						...(reply.date ? { created: reply.date } : {}),
					})),
				],
			};
		});
}

export function openCommentsList(ctx: EditorContext): Promise<string | undefined> {
	const ws = ctx.workbook()?.sheets[ctx.activeSheet()];
	const comments = ws?.comments ?? [];
	const threads = excelCommentThreads(comments);
	let close: (() => void) | undefined;
	let chosen: string | undefined;
	defineCommentsPane(ctx.host.ownerDocument.defaultView?.customElements);
	const pane = ctx.host.ownerDocument.createElement(
		'office-ui-comments-pane',
	) as OfficeUiCommentsPane;
	pane.className = 'xve-comments-list';
	pane.listOnly = true;
	pane.labels = {
		heading: '',
		list: ctx.t('Comments'),
		empty: ctx.t('There are no comments on this sheet.'),
	};
	pane.threads = threads;
	pane.activeThreadId = threads[0]?.id ?? null;
	pane.addEventListener('thread-select', (event) => {
		const id = (event as CustomEvent<{ threadId: string }>).detail.threadId;
		const comment = comments.find((item) => formatAddress(item.address) === id);
		if (!comment) return;
		const at = { ...comment.address };
		chosen = id;
		selectOn(ctx, ctx.activeSheet(), { start: at, end: at });
		close?.();
	});
	return showDialog<string>(
		ctx,
		{
			name: 'comments-list',
			heading: 'Comments',
			okLabel: null,
			body: [pane],
			opened: () => pane.focus(),
		},
		(open) => {
			close = open.close;
		},
	).then((result) => result ?? chosen);
}
