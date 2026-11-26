// Review tab: Workbook Statistics, comments (new / edit, delete, previous, next, show all) and
// Protect Sheet / Protect Workbook.
import { commentsPaneOpen, toggleCommentsPane } from '../comments-pane';
import { allNotesShown, isNote, showNavigatedNote, toggleAllNotes } from '../grid/note-popups';
import type { Comment } from 'ooxml-core/xlsx';
import type { Command } from 'ooxml-core/xlsx/ui';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { toggleWorkbookProtection, unprotectSheet } from '../dialogs/protect-sheet';
import { icon } from 'ooxml-core/xlsx/ui';
import { editing, target, viewing } from 'ooxml-core/xlsx/ui';

const commentAt = (ctx: EditorContext): Comment | undefined => {
	const t = target(ctx);
	return t?.ws.comments.find(
		(c) => c.address.row === t.active.row && c.address.col === t.active.col,
	);
};

const compareAddress = (a: Comment['address'], b: Comment['address']) =>
	a.row - b.row || a.col - b.col;

function step(ctx: EditorContext, delta: 1 | -1, notesOnly = false): void {
	const t = target(ctx);
	if (!t) return;
	const list = t.ws.comments
		.filter((comment) => !notesOnly || isNote(comment))
		.sort((a, b) => delta * compareAddress(a.address, b.address));
	const next =
		list.find((comment) => delta * compareAddress(comment.address, t.active) > 0) ?? list[0];
	if (!next) return;
	const at = { ...next.address };
	ctx.selection.set({ active: at, anchor: at, ranges: [{ start: at, end: at }] });
	ctx.grid()?.scrollTo(at);
	if (notesOnly) showNavigatedNote(ctx, t.ws, at);
}

const hasNotes = (ctx: EditorContext) => target(ctx)?.ws.comments.some(isNote) ?? false;

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
			enabled: (ctx) => !!target(ctx)?.ws.comments.length,
			run: (ctx) => step(ctx, -1),
		}),
		viewing({
			id: 'review.next-comment',
			label: 'Next Comment',
			icon: icon('next'),
			enabled: (ctx) => !!target(ctx)?.ws.comments.length,
			run: (ctx) => step(ctx, 1),
		}),
		viewing({
			id: 'review.show-comments',
			label: 'Show Comments',
			icon: icon('commentShow'),
			checked: commentsPaneOpen,
			run: (ctx, opener) =>
				toggleCommentsPane(ctx, opener instanceof HTMLElement ? opener : undefined),
		}),
		viewing({
			id: 'review.previous-note',
			label: 'Previous Note',
			icon: icon('previous'),
			enabled: hasNotes,
			run: (ctx) => step(ctx, -1, true),
		}),
		viewing({
			id: 'review.next-note',
			label: 'Next Note',
			icon: icon('next'),
			enabled: hasNotes,
			run: (ctx) => step(ctx, 1, true),
		}),
		viewing({
			id: 'review.show-all-notes',
			label: 'Show All Notes',
			icon: icon('commentShow'),
			enabled: (ctx) => allNotesShown(ctx) || hasNotes(ctx),
			checked: allNotesShown,
			run: toggleAllNotes,
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
