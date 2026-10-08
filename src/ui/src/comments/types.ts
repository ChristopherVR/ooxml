// The format-neutral comment model the shared comments pane draws. Products map their own
// comment records (Word `w:comment`, Excel notes and threaded comments, PowerPoint `p:cm`)
// onto it in a small adapter beside their pane; nothing here knows a file format.

/** One comment of a thread: the opening comment or a reply. */
export interface OfficeComment {
	id: string;
	author: string;
	/** Avatar initials; derived from `author` when absent. */
	initials?: string;
	/** Avatar colour (any CSS colour); the accent colour when absent. */
	color?: string;
	/** When it was written: an ISO 8601 date-time is formatted, any other text is shown as is. */
	created?: string;
	/** Plain text; `@mentions` stay text. */
	text: string;
	edited?: boolean;
}

/** A conversation anchored to one place in the document: its first comment and the replies. */
export interface OfficeCommentThread {
	id: string;
	/** Where the thread is anchored, shown above it (a cell address, a slide, a quote). */
	anchorLabel?: string;
	resolved: boolean;
	/** Oldest first; the first entry opens the thread. */
	comments: OfficeComment[];
}

/** Every string the pane shows or announces. The shared package has no dictionary. */
export interface OfficeCommentsLabels {
	/** Pane heading; empty hides it (for a pane inside a titled dialog). */
	heading: string;
	/** Accessible name of the thread list. */
	list: string;
	/** The close button (shown with `closable`). */
	close: string;
	/** Label and placeholder of the new-comment box. */
	newComment: string;
	/** The button that posts a new comment. */
	add: string;
	/** Label and placeholder of a thread's reply box, and its button. */
	reply: string;
	resolve: string;
	reopen: string;
	/** Badge on a resolved thread; empty hides it. */
	resolved: string;
	edit: string;
	save: string;
	cancel: string;
	delete: string;
	/** Marker after an edited comment's time. */
	edited: string;
	/** Shown when there are no threads. */
	empty: string;
}

export const DEFAULT_COMMENTS_LABELS: OfficeCommentsLabels = {
	heading: 'Comments',
	list: 'Comments',
	close: 'Close comments',
	newComment: 'New comment',
	add: 'Add comment',
	reply: 'Reply',
	resolve: 'Resolve',
	reopen: 'Reopen',
	resolved: 'Resolved',
	edit: 'Edit',
	save: 'Save',
	cancel: 'Cancel',
	delete: 'Delete',
	edited: 'Edited',
	empty: 'No comments',
};

/** Product hook classes added to the pane's parts, so a product keeps its own selectors. */
export interface OfficeCommentsClassNames {
	thread?: string;
	comment?: string;
}

export interface OfficeCommentAddDetail {
	text: string;
}
export interface OfficeCommentReplyDetail {
	threadId: string;
	text: string;
}
export interface OfficeCommentEditDetail {
	threadId: string;
	commentId: string;
	text: string;
}
export interface OfficeCommentDeleteDetail {
	threadId: string;
	commentId: string;
}
export interface OfficeThreadDetail {
	threadId: string;
}

/** Up to two initials from an author's name ("Ada Lovelace" is "AL", "grace" is "G"). */
export function commentInitials(author: string): string {
	const words = author.trim().split(/\s+/).filter(Boolean);
	const letters = words.length > 1 ? [words[0]!, words.at(-1)!] : words;
	return letters
		.map((word) => [...word][0] ?? '')
		.join('')
		.toLocaleUpperCase();
}

/** A comment's time for display: ISO date-times are formatted for `locale`, other text is kept. */
export function formatCommentTime(created: string | undefined, locale?: string): string {
	if (!created) return '';
	if (!/^\d{4}-\d{2}-\d{2}/.test(created)) return created;
	const date = new Date(created);
	if (Number.isNaN(date.getTime())) return created;
	try {
		return new Intl.DateTimeFormat(locale, { dateStyle: 'medium', timeStyle: 'short' }).format(
			date,
		);
	} catch {
		return created;
	}
}
