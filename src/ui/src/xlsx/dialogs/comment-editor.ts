import { target, type EditorContext } from 'ooxml-core/xlsx/ui';
import { checkbox, field, invalid, textArea } from './fields';
import { showDialog } from './frame';

/** The existing comment command also edits the replies of an imported thread. */
export function openCommentEditor(ctx: EditorContext): Promise<string | undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const at = t.active;
	const existing = t.ws.comments.find((c) => c.address.row === at.row && c.address.col === at.col);
	const initialText = existing?.text;
	const initialXml = existing?.threadedXml;
	const initialResolved = existing?.resolved ?? false;
	const root = textArea(ctx, initialText ?? '', 4);
	const replies = (existing?.replies ?? []).map((reply, index) => ({
		source: reply,
		text: reply.text,
		xml: reply.threadedXml,
		author: reply.author,
		index,
		input: textArea(ctx, reply.text),
	}));
	const resolved = existing?.replies
		? checkbox(ctx, 'Resolved', existing.resolved ?? false)
		: undefined;
	const body = [
		field(ctx, 'Comment', root),
		...replies.map(({ input, author }, index) =>
			field(
				ctx,
				ctx.t('Reply {number} ({author})', {
					number: index + 1,
					author,
				}),
				input,
			),
		),
	];
	if (resolved) body.push(resolved.wrapper);
	return showDialog<string>(ctx, {
		name: 'comment',
		heading: existing ? 'Edit Comment' : 'New Comment',
		body,
		opened: () => {
			root.focus();
			root.select();
		},
		submit: () => {
			const rootChanged = !existing || root.value !== initialText;
			const changed = replies.filter((reply) => reply.input.value !== reply.text);
			const resolutionChanged = resolved && resolved.input.checked !== initialResolved;
			for (const input of [
				...(rootChanged ? [root] : []),
				...changed.map((reply) => reply.input),
			]) {
				if (!input.value.trim()) return invalid(ctx, input, 'Type a comment first.');
			}
			if (!rootChanged && !changed.length && !resolutionChanged) return root.value;
			const sheet = t.session.workbook.sheets.indexOf(t.ws);
			const current = t.ws.comments.find(
				(c) => c.address.row === at.row && c.address.col === at.col,
			);
			const conflict =
				ctx.session() !== t.session ||
				sheet < 0 ||
				!!existing !== !!current ||
				initialXml !== current?.threadedXml ||
				(rootChanged && current && current.text !== initialText && current.text !== root.value) ||
				changed.some((reply) => {
					const now = current?.replies?.[reply.index];
					return (
						!now ||
						(reply.xml !== undefined ? reply.xml !== now.threadedXml : reply.source !== now) ||
						(now.text !== reply.text && now.text !== reply.input.value)
					);
				});
			if (conflict) return invalid(ctx, root, 'The comment changed. Reopen it and try again.');
			t.session.batch(existing ? 'Edit comment' : 'New comment', () => {
				if (rootChanged && root.value !== current?.text) {
					t.session.setComment(sheet, at, root.value, current?.author ?? ctx.authorName());
				}
				changed.forEach(({ input, index }) => {
					t.session.setCommentReply(sheet, at, index, input.value);
				});
				if (resolutionChanged) t.session.setCommentResolved(sheet, at, resolved.input.checked);
			});
			return root.value;
		},
	});
}
