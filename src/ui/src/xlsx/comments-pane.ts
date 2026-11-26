import { formatAddress, parseAddress, type Comment } from 'ooxml-core/xlsx';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import type { OfficeUiCommentsPane } from '../comments/comments-pane';
import type { OfficeCommentThread } from '../comments/types';
import { modalActiveElement } from '../dialog/focus';

const panes = new WeakMap<
	EditorContext,
	{ element: HTMLElement; toggle(opener: Element | null): Promise<void> }
>();

export function commentsPaneOpen(ctx: EditorContext): boolean {
	return panes.get(ctx)?.element.hidden === false;
}

export async function toggleCommentsPane(
	ctx: EditorContext,
	opener = modalActiveElement(ctx.root),
): Promise<void> {
	await panes.get(ctx)?.toggle(opener);
}

/** Excel threaded comments as neutral threads, in sheet order, keyed by address. */
function excelCommentThreads(comments: readonly Comment[]): OfficeCommentThread[] {
	return comments
		.filter((comment) => comment.replies !== undefined)
		.sort((a, b) => a.address.row - b.address.row || a.address.col - b.address.col)
		.map((comment) => {
			const id = formatAddress(comment.address);
			return {
				id,
				anchorLabel: id,
				resolved: comment.resolved ?? false,
				comments: [
					{
						id: `${id}:0`,
						author: comment.author,
						text: comment.text,
						...(comment.date ? { created: comment.date } : {}),
					},
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

export function createCommentsPane(ctx: EditorContext) {
	const doc = ctx.host.ownerDocument;
	const element = doc.createElement('aside');
	element.className = 'xve-comments-pane';
	element.hidden = true;
	element.tabIndex = -1;
	const heading = doc.createElement('h2');
	const close = doc.createElement('button');
	close.type = 'button';
	close.className = 'xve-icon-button';
	close.textContent = '×';
	const header = doc.createElement('header');
	header.append(heading, close);
	const add = doc.createElement('button');
	add.type = 'button';
	add.className = 'xve-comments-new';
	element.append(header, add);

	let pane: OfficeUiCommentsPane | undefined;
	let opener: Element | null = null;
	let pending = false;
	let source: readonly Comment[] | undefined;
	let threadsDirty = true;
	let locale: string | undefined;

	const refresh = () => {
		if (element.hidden || !pane) return;
		const sheet = ctx.workbook()?.sheets[ctx.activeSheet()];
		const active = ctx.selection.get().active;
		pane.activeThreadId = formatAddress(active);
		add.disabled = !ctx.commands.isEnabled('review.new-comment');
		add.title = add.disabled
			? ctx.t(
					!sheet
						? 'No workbook is open.'
						: ctx.readOnly()
							? 'Read-only'
							: 'Editing comments is unavailable.',
				)
			: '';
		const liveLocale = ctx.locale();
		if (locale !== liveLocale) {
			locale = liveLocale;
			heading.textContent = ctx.t('Comments');
			element.setAttribute('aria-label', heading.textContent);
			close.setAttribute('aria-label', ctx.t('Close'));
			add.textContent = ctx.t('New Comment');
			pane.locale = liveLocale;
			pane.labels = {
				heading: '',
				list: heading.textContent,
				resolved: ctx.t('Resolved'),
				empty: ctx.t('There are no comments on this sheet.'),
			};
		}
		if (threadsDirty || source !== sheet?.comments) {
			source = sheet?.comments;
			pane.threads = excelCommentThreads(source ?? []);
			threadsDirty = false;
		}
	};
	ctx.onModelChange((change) => {
		if (element.hidden || !change || typeof change !== 'object' || !('kind' in change)) return;
		// Comment edits and shared updates can mutate an existing array. Cell edits that
		// replace annotations change its identity, which refresh also checks.
		if (
			['annotations', 'batch', 'undo', 'redo', 'remote', 'load'].some(
				(kind) => kind === change.kind,
			)
		)
			threadsDirty = true;
	});

	const dismiss = () => {
		element.hidden = true;
		ctx.requestRender();
		if (opener instanceof HTMLElement && opener.isConnected) opener.focus({ preventScroll: true });
		else ctx.grid()?.focus();
		opener = null;
	};
	close.addEventListener('click', dismiss);
	add.addEventListener('click', () => void ctx.commands.run('review.new-comment'));
	element.addEventListener('keydown', (event) => {
		if (event.key !== 'Escape' || event.defaultPrevented) return;
		event.preventDefault();
		event.stopPropagation();
		dismiss();
	});
	panes.set(ctx, {
		element,
		async toggle(trigger) {
			if (!element.hidden) {
				dismiss();
				return;
			}
			if (pending) return;
			opener = trigger;
			if (!pane) {
				pending = true;
				try {
					const { defineCommentsPane, OfficeUiCommentsPane } =
						await import('../comments/comments-pane');
					if (!element.isConnected) return;
					defineCommentsPane(doc.defaultView?.customElements);
					const list = doc.createElement('office-ui-comments-pane');
					if (!(list instanceof OfficeUiCommentsPane)) return;
					pane = list;
					pane.listOnly = true;
					pane.addEventListener('thread-select', (event) => {
						if (!(event instanceof CustomEvent) || typeof event.detail?.threadId !== 'string')
							return;
						const active = parseAddress(event.detail.threadId);
						if (!active) return;
						ctx.selection.set({ active, anchor: active, ranges: [{ start: active, end: active }] });
						ctx.grid()?.scrollTo(active);
					});
					element.append(pane);
				} finally {
					pending = false;
				}
			}
			element.hidden = false;
			threadsDirty = true;
			refresh();
			(add.disabled ? close : add).focus();
			ctx.requestRender();
		},
	});
	return { element, refresh };
}
