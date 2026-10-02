// Review tab: Workbook Statistics, comments (new / edit, delete, previous, next, show all) and
// Protect Sheet / Protect Workbook.
import type { Comment } from '@christophervr/xlsx-core';
import type { Command } from '../commands.js';
import type { EditorContext } from '../context.js';
import { toggleWorkbookProtection, unprotectSheet } from '../dialogs/protect-sheet.js';
import { icon } from './icons.js';
import { editing, target, viewing } from './util.js';

const commentAt = (ctx: EditorContext): Comment | undefined => {
	const t = target(ctx);
	return t?.ws.comments.find(
		(c) => c.address.row === t.active.row && c.address.col === t.active.col,
	);
};

const ordered = (ctx: EditorContext): Comment[] =>
	[...(target(ctx)?.ws.comments ?? [])].sort(
		(a, b) => a.address.row - b.address.row || a.address.col - b.address.col,
	);

function step(ctx: EditorContext, delta: 1 | -1): void {
	const t = target(ctx);
	const list = ordered(ctx);
	if (!t || !list.length) return;
	const key = (c: Comment) => c.address.row * 20000 + c.address.col;
	const here = t.active.row * 20000 + t.active.col;
	const next =
		delta > 0
			? (list.find((c) => key(c) > here) ?? list[0])
			: ([...list].reverse().find((c) => key(c) < here) ?? list[list.length - 1]);
	if (!next) return;
	const at = { ...next.address };
	ctx.selection.set({ active: at, anchor: at, ranges: [{ start: at, end: at }] });
	ctx.grid()?.scrollTo(at);
}

export function reviewCommands(): Command[] {
	return [
		viewing({
			id: 'review.statistics',
			label: 'Workbook Statistics',
			icon: icon('statistics'),
			run: (ctx) => void ctx.dialogs.open('workbook-statistics'),
		}),
		editing({
			id: 'review.new-comment',
			label: 'New Comment',
			icon: icon('commentNew'),
			shortcut: 'Shift+F2',
			lock: 'objects',
			run: (ctx) => void ctx.dialogs.open('comment'),
		}),
		editing({
			id: 'review.delete-comment',
			label: 'Delete Comment',
			icon: icon('commentDelete'),
			lock: 'objects',
			enabled: (ctx) => {
				const t = target(ctx);
				return (
					!!t &&
					t.ws.comments.some((c) =>
						t.ranges.some(
							(r) =>
								c.address.row >= r.start.row &&
								c.address.row <= r.end.row &&
								c.address.col >= r.start.col &&
								c.address.col <= r.end.col,
						),
					)
				);
			},
			run: (ctx) => {
				const t = target(ctx);
				if (!t) return;
				const doomed = t.ws.comments.filter((c) =>
					t.ranges.some(
						(r) =>
							c.address.row >= r.start.row &&
							c.address.row <= r.end.row &&
							c.address.col >= r.start.col &&
							c.address.col <= r.end.col,
					),
				);
				t.session.batch('Delete comment', () => {
					for (const c of doomed)
						t.session.setComment(t.sheet, c.address, undefined, ctx.authorName());
				});
			},
		}),
		viewing({
			id: 'review.previous-comment',
			label: 'Previous Comment',
			icon: icon('previous'),
			enabled: (ctx) => ordered(ctx).length > 0,
			run: (ctx) => step(ctx, -1),
		}),
		viewing({
			id: 'review.next-comment',
			label: 'Next Comment',
			icon: icon('next'),
			enabled: (ctx) => ordered(ctx).length > 0,
			run: (ctx) => step(ctx, 1),
		}),
		viewing({
			id: 'review.show-comments',
			label: 'Show Comments',
			icon: icon('commentShow'),
			run: (ctx) => void ctx.dialogs.open('comments-list'),
		}),
		editing({
			id: 'review.protect-sheet',
			label: 'Protect Sheet',
			icon: icon('protectSheet'),
			lock: false,
			checked: (ctx) => !!target(ctx)?.ws.protection?.sheet,
			run: async (ctx) => {
				const t = target(ctx);
				if (!t) return;
				if (!t.ws.protection?.sheet) return void ctx.dialogs.open('protect-sheet');
				await unprotectSheet(ctx);
			},
		}),
		editing({
			id: 'review.protect-workbook',
			label: 'Protect Workbook',
			icon: icon('protectWorkbook'),
			lock: false,
			checked: (ctx) => ctx.workbook()?.structureLocked === true,
			run: async (ctx) => {
				if (!(await toggleWorkbookProtection(ctx))) return;
				ctx.toast(
					ctx.t(
						ctx.workbook()?.structureLocked
							? 'Workbook structure is protected.'
							: 'Workbook structure is no longer protected.',
					),
					'info',
				);
			},
		}),
	];
}

export { commentAt };
