import { target, type EditorContext } from 'ooxml-core/xlsx/ui';
import { field, invalid, textArea } from './fields';
import { showDialog } from './frame';

/** The existing comment command also edits the replies of an imported thread. */
export function openCommentEditor(ctx: EditorContext): Promise<string | undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const at = t.active;
	const existing = t.ws.comments.find((c) => c.address.row === at.row && c.address.col === at.col);
	const root = textArea(ctx, existing?.text ?? '', 4);
	const replies = (existing?.replies ?? []).map((reply) => textArea(ctx, reply.text));
	const body = [
		field(ctx, 'Comment', root),
		...replies.map((input, index) =>
			field(
				ctx,
				ctx.t('Reply {number} ({author})', {
					number: index + 1,
					author: existing!.replies![index]!.author,
				}),
				input,
			),
		),
	];
	return showDialog<string>(ctx, {
		name: 'comment',
		heading: existing ? 'Edit Comment' : 'New Comment',
		body,
		opened: () => {
			root.focus();
			root.select();
		},
		submit: () => {
			for (const input of [root, ...replies]) {
				if (!input.value.trim()) return invalid(ctx, input, 'Type a comment first.');
			}
			t.session.batch(existing ? 'Edit comment' : 'New comment', () => {
				if (!existing || root.value !== existing.text) {
					t.session.setComment(t.sheet, at, root.value, existing?.author ?? ctx.authorName());
				}
				replies.forEach((input, index) => {
					t.session.setCommentReply(t.sheet, at, index, input.value);
				});
			});
			return root.value;
		},
	});
}
