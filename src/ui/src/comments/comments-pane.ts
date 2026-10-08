import { html, nothing, type TemplateResult } from 'lit';
import { ifDefined } from 'lit/directives/if-defined.js';
import { repeat } from 'lit/directives/repeat.js';
import { OfficeElement, controlStyles, flag } from '../base';
import { definer, present } from '../registry';
import css from './comments-pane.css?raw';
import {
	DEFAULT_COMMENTS_LABELS,
	commentInitials,
	formatCommentTime,
	type OfficeComment,
	type OfficeCommentThread,
	type OfficeCommentsClassNames,
	type OfficeCommentsLabels,
} from './types';

const INTERACTIVE = new Set(['BUTTON', 'TEXTAREA', 'INPUT', 'A', 'SELECT']);

/**
 * `<office-ui-comments-pane>`: Office's comments pane, one UX for Word, Excel and PowerPoint.
 * Threads show each comment's avatar initials, author, time and text, with Reply,
 * Resolve/Reopen, Edit (own comments, by `currentAuthor`) and Delete, under a new-comment box.
 * The pane never changes its threads: every action is an event and the product sets `threads`
 * again. Arrow keys, Home and End move between threads; Enter or Space (or a click) selects one
 * (`thread-select`) so the product can go to its anchor. Escape is left to the host.
 *
 * Properties: `threads`, `activeThreadId`, `readOnly` (controls shown but disabled),
 * `listOnly` (no composer, replies or actions: a navigation list), `closable`, `addDisabled`,
 * `currentAuthor`, `labels` (partial, English defaults), `classNames` (product hook classes),
 * `locale` (time formatting). Events (bubbling, composed): `comment-add` `{ text }`,
 * `comment-reply` `{ threadId, text }`, `comment-edit` `{ threadId, commentId, text }`,
 * `comment-delete` `{ threadId, commentId }`, `thread-resolve`, `thread-reopen` and
 * `thread-select` `{ threadId }`, `comments-close`.
 */
export class OfficeUiCommentsPane extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		threads: { attribute: false },
		activeThreadId: { attribute: 'active-thread-id', type: String },
		readOnly: { attribute: 'read-only', ...flag },
		listOnly: { attribute: 'list-only', ...flag },
		closable: flag,
		addDisabled: { attribute: 'add-disabled', ...flag },
		currentAuthor: { attribute: 'current-author', type: String },
		labels: { attribute: false },
		classNames: { attribute: false },
		locale: { type: String },
	};
	declare threads: readonly OfficeCommentThread[];
	declare activeThreadId: string | null;
	declare readOnly: boolean;
	declare listOnly: boolean;
	declare closable: boolean;
	declare addDisabled: boolean;
	declare currentAuthor: string | null;
	declare labels: Partial<OfficeCommentsLabels>;
	declare classNames: OfficeCommentsClassNames;
	declare locale: string | null;
	private focusedId: string | undefined;
	private editingId: string | undefined;

	constructor() {
		super();
		this.threads = [];
		this.activeThreadId = null;
		this.readOnly = false;
		this.listOnly = false;
		this.closable = false;
		this.addDisabled = false;
		this.currentAuthor = null;
		this.labels = {};
		this.classNames = {};
		this.locale = null;
	}

	private get text(): OfficeCommentsLabels {
		return { ...DEFAULT_COMMENTS_LABELS, ...this.labels };
	}

	private get locked(): boolean {
		return present(this.readOnly);
	}

	/** The thread that takes Tab focus: the last one focused, else the active one, else the first. */
	private get rovingId(): string | undefined {
		const ids = this.threads.map((thread) => thread.id);
		if (this.focusedId && ids.includes(this.focusedId)) return this.focusedId;
		if (this.activeThreadId && ids.includes(this.activeThreadId)) return this.activeThreadId;
		return ids[0];
	}

	private card(id: string | undefined): HTMLElement | null {
		if (id === undefined) return null;
		return (
			[...this.renderRoot.querySelectorAll<HTMLElement>('.thread')].find(
				(card) => card.dataset.threadId === id,
			) ?? null
		);
	}

	/** Focus the current thread, or the new-comment box when there is none. */
	override focus(options?: FocusOptions): void {
		const target =
			this.card(this.rovingId) ?? this.renderRoot.querySelector<HTMLElement>('textarea');
		if (target) target.focus(options);
		else super.focus(options);
	}

	private submitNew(): void {
		const box = this.renderRoot.querySelector<HTMLTextAreaElement>('textarea.new');
		const text = box?.value.trim();
		if (!box || !text || this.locked || present(this.addDisabled)) return;
		box.value = '';
		this.fire('comment-add', { text });
	}

	private submitReply(threadId: string): void {
		const box = this.card(threadId)?.querySelector<HTMLTextAreaElement>('textarea.reply-box');
		const text = box?.value.trim();
		if (!box || !text || this.locked) return;
		box.value = '';
		this.fire('comment-reply', { threadId, text });
	}

	private startEdit(commentId: string): void {
		this.editingId = commentId;
		this.requestUpdate();
		this.renderRoot.querySelector<HTMLTextAreaElement>('textarea.edit-box')?.focus();
	}

	private stopEdit(): void {
		this.editingId = undefined;
		this.requestUpdate();
	}

	private saveEdit(thread: OfficeCommentThread, comment: OfficeComment): void {
		const box = this.renderRoot.querySelector<HTMLTextAreaElement>('textarea.edit-box');
		const text = box?.value.trim();
		this.stopEdit();
		if (!text || text === comment.text || this.locked) return;
		this.fire('comment-edit', { threadId: thread.id, commentId: comment.id, text });
	}

	/** Ctrl+Enter (Cmd+Enter) posts from any box, like Office. */
	private onBoxKey(event: KeyboardEvent, submit: () => void): void {
		if (event.key === 'Enter' && (event.ctrlKey || event.metaKey) && !event.isComposing) {
			event.preventDefault();
			submit();
		}
	}

	private onEditKey(event: KeyboardEvent, thread: OfficeCommentThread, comment: OfficeComment) {
		if (event.key !== 'Escape') return this.onBoxKey(event, () => this.saveEdit(thread, comment));
		// Escape cancels the edit only; the host keeps its own Escape for the pane.
		event.preventDefault();
		event.stopPropagation();
		this.stopEdit();
	}

	private select(threadId: string): void {
		this.focusedId = threadId;
		this.requestUpdate();
		this.fire('thread-select', { threadId });
	}

	private onThreadClick(event: MouseEvent, threadId: string): void {
		const inner = event
			.composedPath()
			.some((node) => node instanceof HTMLElement && INTERACTIVE.has(node.tagName));
		if (!inner) this.select(threadId);
	}

	private onListKey(event: KeyboardEvent): void {
		const card = event.composedPath()[0];
		if (!(card instanceof HTMLElement) || !card.classList.contains('thread')) return;
		const ids = this.threads.map((thread) => thread.id);
		const at = ids.indexOf(card.dataset.threadId ?? '');
		if (at < 0) return;
		const moves: Record<string, number> = {
			ArrowDown: Math.min(at + 1, ids.length - 1),
			ArrowUp: Math.max(at - 1, 0),
			Home: 0,
			End: ids.length - 1,
		};
		const move = moves[event.key];
		if (move !== undefined) {
			event.preventDefault();
			this.focusedId = ids[move];
			this.requestUpdate();
			this.card(this.focusedId)?.focus();
		} else if (event.key === 'Enter' || event.key === ' ') {
			event.preventDefault();
			this.select(ids[at]!);
		}
	}

	private renderActions(thread: OfficeCommentThread, comment: OfficeComment, first: boolean) {
		if (present(this.listOnly)) return nothing;
		const l = this.text;
		const locked = this.locked;
		const own = Boolean(this.currentAuthor) && comment.author === this.currentAuthor;
		const threadId = thread.id;
		const status = thread.resolved ? 'reopen' : 'resolve';
		return html`<div class="actions"
			>${first ? html`<button type="button" data-action=${status} ?disabled=${locked} @click=${() => (thread.resolved ? this.fire('thread-reopen', { threadId }) : this.fire('thread-resolve', { threadId }))}>${thread.resolved ? l.reopen : l.resolve}</button>` : nothing}${own && this.editingId !== comment.id ? html`<button type="button" data-action="edit" ?disabled=${locked} @click=${() => this.startEdit(comment.id)}>${l.edit}</button>` : nothing}<button
				type="button"
				data-action="delete"
				?disabled=${locked}
				@click=${() => this.fire('comment-delete', { threadId, commentId: comment.id })}
				>${l.delete}</button
			></div
		>`;
	}

	private renderComment(
		thread: OfficeCommentThread,
		comment: OfficeComment,
		first: boolean,
	): TemplateResult {
		const l = this.text;
		const time = formatCommentTime(comment.created, this.locale ?? undefined);
		const body =
			this.editingId === comment.id && !this.locked
				? html`<div class="edit">
						<textarea
							class="edit-box"
							aria-label=${l.edit}
							.value=${comment.text}
							@keydown=${(event: KeyboardEvent) => this.onEditKey(event, thread, comment)}
						></textarea
						><div class="actions"
							><button
								type="button"
								data-action="save"
								@click=${() => this.saveEdit(thread, comment)}
								>${l.save}</button
							><button type="button" data-action="cancel" @click=${() => this.stopEdit()}
								>${l.cancel}</button
							></div
						></div
					>`
				: html`<p class="text">${comment.text}</p>`;
		return html`<div
			class="comment ${first ? 'root' : 'reply'} ${this.classNames.comment ?? ''}"
			data-comment-id=${comment.id}
			data-resolved=${String(thread.resolved)}
			><div class="meta"
				><span
					class="avatar"
					aria-hidden="true"
					style=${ifDefined(comment.color ? `background:${comment.color}` : undefined)}
					>${comment.initials || commentInitials(comment.author)}</span
				><span class="author">${comment.author}</span
				>${time ? html`<time datetime=${ifDefined(comment.created)}>${time}</time>` : nothing}${comment.edited ? html`<span class="edited">${l.edited}</span>` : nothing}</div
			>${body}${this.renderActions(thread, comment, first)}</div
		>`;
	}

	private renderThread(thread: OfficeCommentThread, roving: string | undefined): TemplateResult {
		const l = this.text;
		const [first, ...replies] = thread.comments;
		const badge = thread.resolved && l.resolved;
		const anchor =
			thread.anchorLabel || badge
				? html`<div class="anchor"
						>${thread.anchorLabel ? html`<span class="anchor-label">${thread.anchorLabel}</span>` : nothing}${badge ? html`<span class="badge">${l.resolved}</span>` : nothing}</div
					>`
				: nothing;
		const reply = present(this.listOnly)
			? nothing
			: html`<div class="reply">
					<textarea
						class="reply-box"
						aria-label=${l.reply}
						placeholder=${l.reply}
						?disabled=${this.locked}
						@keydown=${(event: KeyboardEvent) => this.onBoxKey(event, () => this.submitReply(thread.id))}
					></textarea
					><button
						type="button"
						data-action="reply"
						?disabled=${this.locked}
						@click=${() => this.submitReply(thread.id)}
						>${l.reply}</button
					></div
				>`;
		return html`<article
			role="listitem"
			class="thread ${this.classNames.thread ?? ''}"
			data-thread-id=${thread.id}
			data-resolved=${String(thread.resolved)}
			aria-current=${ifDefined(thread.id === this.activeThreadId ? 'true' : undefined)}
			tabindex=${thread.id === roving ? '0' : '-1'}
			@click=${(event: MouseEvent) => this.onThreadClick(event, thread.id)}
			>${anchor}${first ? this.renderComment(thread, first, true) : nothing}${replies.map((comment) => this.renderComment(thread, comment, false))}${reply}</article
		>`;
	}

	protected override render() {
		const l = this.text;
		const closable = present(this.closable);
		const header =
			l.heading || closable
				? html`<div class="header"
						>${l.heading ? html`<h2>${l.heading}</h2>` : nothing}${closable ? html`<button type="button" class="close" data-action="close" aria-label=${l.close} @click=${() => this.fire('comments-close', {})}>${l.close}</button>` : nothing}</div
					>`
				: nothing;
		const composer = present(this.listOnly)
			? nothing
			: html`<div class="composer">
					<textarea
						class="new"
						aria-label=${l.newComment}
						placeholder=${l.newComment}
						?disabled=${this.locked}
						@keydown=${(event: KeyboardEvent) => this.onBoxKey(event, () => this.submitNew())}
					></textarea
					><button
						type="button"
						data-action="add"
						?disabled=${this.locked || present(this.addDisabled)}
						@click=${() => this.submitNew()}
						>${l.add}</button
					></div
				>`;
		const roving = this.rovingId;
		const list = this.threads.length
			? html`<div class="threads" role="list" aria-label=${l.list} @keydown=${this.onListKey}
					>${repeat(
						this.threads,
						(thread) => thread.id,
						(thread) => this.renderThread(thread, roving),
					)}</div
				>`
			: html`<p class="empty">${l.empty}</p>`;
		return html`${header}${composer}${list}`;
	}
}

export const defineCommentsPane = definer('office-ui-comments-pane', () => OfficeUiCommentsPane);
