// Word's comments pane: the shared `office-ui-comments-pane` fed from the document's comments.
import type { Comment, DocumentModel } from 'ooxml-core/docx';
import { defineCommentsPane, type OfficeUiCommentsPane } from '../comments/comments-pane';
import type { OfficeCommentThread, OfficeCommentsLabels } from '../comments/types';
import { normalizeEditorLocale, translate, type EditorLocale } from './localization';

export interface CommentsPanelOptions {
	getModel: () => DocumentModel;
	canAdd: () => boolean;
	canEdit?: () => boolean;
	/** The thread anchored at the caret, highlighted in the pane. */
	activeId?: () => string | undefined;
	onAdd: (text: string) => void;
	onReply: (parentId: string, text: string) => void;
	onResolve: (id: string, resolved: boolean) => void;
	onDelete: (id: string) => void;
	/** A thread was chosen in the pane: go to its anchor. */
	onSelect?: (id: string) => void;
	onClose?: () => void;
}
export interface CommentsPanelHandle {
	element: HTMLElement;
	open(): void;
	close(): void;
	refresh(): void;
	setLocale(locale: string): void;
	readonly isOpen: boolean;
}

/** Word comments as neutral threads: each top-level comment with its direct replies. */
export function wordCommentThreads(comments: readonly Comment[]): OfficeCommentThread[] {
	const entry = (comment: Comment) => ({
		id: comment.id,
		author: comment.author,
		text: comment.text,
		...(comment.initials ? { initials: comment.initials } : {}),
		...(comment.date ? { created: comment.date } : {}),
	});
	return comments
		.filter((comment) => !comment.parentId)
		.map((comment) => ({
			id: comment.id,
			resolved: Boolean(comment.resolved),
			comments: [
				entry(comment),
				...comments.filter((reply) => reply.parentId === comment.id).map(entry),
			],
		}));
}

function labels(locale: EditorLocale): Partial<OfficeCommentsLabels> {
	const t = (key: Parameters<typeof translate>[1]) => translate(locale, key);
	return {
		heading: t('Comments'),
		list: t('Comments'),
		close: t('Close comments'),
		newComment: t('New comment'),
		add: t('Add comment'),
		reply: t('Reply'),
		resolve: t('Resolve'),
		reopen: t('Reopen'),
		resolved: t('Resolved'),
		delete: t('Delete'),
		empty: t('No comments'),
	};
}

type Detail = { text: string; threadId: string; commentId: string };
const detail = (event: Event) => (event as CustomEvent<Detail>).detail;

export function createCommentsPanel(options: CommentsPanelOptions): CommentsPanelHandle {
	let locale: EditorLocale = 'en';
	defineCommentsPane();
	const panel = document.createElement('office-ui-comments-pane') as OfficeUiCommentsPane;
	panel.className = 'dve-comments-panel';
	panel.setAttribute('role', 'complementary');
	panel.setAttribute('aria-label', translate(locale, 'Comments'));
	panel.hidden = true;
	panel.closable = true;
	// The e2e specs find a comment and its own buttons by `.dve-comment-thread`.
	panel.classNames = { thread: 'dve-comment-card', comment: 'dve-comment-thread' };
	panel.labels = labels(locale);
	panel.addEventListener('comment-add', (event) => {
		options.onAdd(detail(event).text);
		refresh();
	});
	panel.addEventListener('comment-reply', (event) =>
		options.onReply(detail(event).threadId, detail(event).text),
	);
	panel.addEventListener('thread-resolve', (event) =>
		options.onResolve(detail(event).threadId, true),
	);
	panel.addEventListener('thread-reopen', (event) =>
		options.onResolve(detail(event).threadId, false),
	);
	panel.addEventListener('comment-delete', (event) => options.onDelete(detail(event).commentId));
	panel.addEventListener('thread-select', (event) => options.onSelect?.(detail(event).threadId));
	panel.addEventListener('comments-close', () => close());

	let isOpen = false;
	function close() {
		isOpen = false;
		panel.hidden = true;
		options.onClose?.();
	}
	function refresh() {
		panel.readOnly = !(options.canEdit?.() ?? true);
		panel.addDisabled = !options.canAdd();
		panel.activeThreadId = options.activeId?.() ?? null;
		panel.threads = wordCommentThreads(options.getModel().comments ?? []);
	}
	refresh();
	return {
		element: panel,
		get isOpen() {
			return isOpen;
		},
		open() {
			isOpen = true;
			panel.hidden = false;
			refresh();
		},
		close,
		refresh,
		setLocale(value: string) {
			locale = normalizeEditorLocale(value);
			panel.setAttribute('aria-label', translate(locale, 'Comments'));
			panel.labels = labels(locale);
			panel.locale = locale;
			refresh();
		},
	};
}
