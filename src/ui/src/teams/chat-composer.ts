import { html, nothing, type PropertyValues } from 'lit';
import { live } from 'lit/directives/live.js';
import { definer } from '../registry';
import { TeamsElement, withStyles } from './base';
import { icon } from './icons';
import css from './chat-composer.css?raw';
import {
	MAX_ATTACHMENTS,
	type ChatDraft,
	type MentionQuery,
	type MentionTarget,
	filterMentionTargets,
	insertMention,
	mentionQueryAt,
} from 'ooxml-core/teams';

/** The detail of `office-chat-send`: the draft plus the people picked with @. */
export interface ChatSendDetail {
	text: string;
	files: ChatDraft['files'];
	mentions?: MentionTarget[];
}

export const EMOJI = [
	'😀',
	'😄',
	'😂',
	'🙂',
	'😉',
	'😍',
	'🤔',
	'😮',
	'😢',
	'😡',
	'👍',
	'👎',
	'👏',
	'🙏',
	'💪',
	'🎉',
	'🔥',
	'❤️',
	'✅',
	'👀',
] as const;
const MAX_FILES = MAX_ATTACHMENTS;

/**
 * The compose box: a growing text area, attach and emoji buttons, reply or edit banner and a
 * typing line. Enter sends, Shift+Enter adds a line. Properties: `typing` (names), `replyingTo`
 * (a label or null), `editing` (true while editing), `disabled`, `placeholder`, and `value`.
 * Events: `office-chat-draft` `{ text, files, missingFiles }` after a user change,
 * `office-chat-send` `{ text, files, mentions }`, `office-chat-typing` `{}` (throttled) and
 * `office-chat-cancel` `{}` (the banner's close button or Escape). File bytes are never read here:
 * the host stores them wherever its server says and posts a link.
 *
 * Typing `@` opens a people picker over `mentionTargets` (people, then @channel and @team): it
 * filters as you type, Up / Down move, Enter or Tab insert `@Name`, Escape closes it. The people
 * picked go out as `mentions` with the text.
 */
export class OfficeUiChatComposer extends TeamsElement {
	static override styles = withStyles(css);
	static override properties = {
		typing: { attribute: false },
		replyingTo: { attribute: false },
		editing: { type: Boolean },
		disabled: { type: Boolean, reflect: true },
		placeholder: { type: String },
		value: { type: String },
		files: { state: true },
		missingFiles: { attribute: false },
		emojiOpen: { state: true },
		mentionTargets: { attribute: false },
		picker: { state: true },
		pickerIndex: { state: true },
	};
	declare typing: string[];
	declare replyingTo: string | null;
	declare editing: boolean;
	declare disabled: boolean;
	declare placeholder: string;
	declare value: string;
	declare files: ChatDraft['files'];
	declare missingFiles: string[];
	declare emojiOpen: boolean;
	declare mentionTargets: MentionTarget[];
	declare picker: MentionQuery | null;
	declare pickerIndex: number;
	private inserted: MentionTarget[] = [];
	private lastTyping = 0;

	constructor() {
		super();
		this.typing = [];
		this.replyingTo = null;
		this.editing = false;
		this.disabled = false;
		this.placeholder = 'Type a message';
		this.value = '';
		this.files = [];
		this.missingFiles = [];
		this.emojiOpen = false;
		this.mentionTargets = [];
		this.picker = null;
		this.pickerIndex = 0;
	}

	private focusMessage(): void {
		void this.updateComplete.then(() =>
			this.renderRoot.querySelector('textarea')?.focus({ preventScroll: true }),
		);
	}
	private resizeMessage(): void {
		const area = this.renderRoot.querySelector('textarea');
		if (!area) return;
		area.style.height = 'auto';
		area.style.height = `${Math.min(area.scrollHeight, 160)}px`;
	}
	protected override updated(changed: PropertyValues<this>): void {
		if (changed.has('value')) this.resizeMessage();
		if (changed.has('replyingTo') || changed.has('editing')) {
			if (this.replyingTo !== null || this.editing)
				this.renderRoot.querySelector('textarea')?.focus();
		}
	}

	private get canSend(): boolean {
		return (
			!this.disabled &&
			!this.missingFiles.length &&
			(this.value.trim() !== '' || this.files.length > 0)
		);
	}

	private submit(): void {
		if (!this.canSend) return;
		const { value: text, files, inserted: mentions } = this;
		this.value = '';
		this.files = [];
		this.inserted = [];
		this.emojiOpen = false;
		this.picker = null;
		const detail: ChatSendDetail = { text, files, ...(mentions.length ? { mentions } : {}) };
		this.fire('office-chat-send', detail);
		this.focusMessage();
	}

	private onInput(event: Event): void {
		this.value = (event.target as HTMLTextAreaElement).value;
		this.syncPicker();
		this.saveDraft();
		const now = Date.now();
		if (now - this.lastTyping > 2000) {
			this.lastTyping = now;
			this.fire('office-chat-typing', {});
		}
	}

	/** Open the picker while the caret sits in an `@query` that somebody matches. */
	private syncPicker(): void {
		const area = this.renderRoot.querySelector('textarea');
		const query = area ? mentionQueryAt(this.value, area.selectionStart) : null;
		if (!query || filterMentionTargets(this.mentionTargets, query.query).length === 0) {
			this.picker = null;
			return;
		}
		if (this.picker?.start !== query.start || this.picker.query !== query.query)
			this.pickerIndex = 0;
		this.picker = query;
	}

	private pickerOptions(): MentionTarget[] {
		return this.picker ? filterMentionTargets(this.mentionTargets, this.picker.query) : [];
	}

	private chooseMention(target: MentionTarget): void {
		const area = this.renderRoot.querySelector('textarea');
		const picker = this.picker;
		if (!area || !picker) return;
		const next = insertMention(this.value, picker, area.selectionStart, target);
		this.value = next.text;
		this.inserted = [
			...this.inserted.filter((t) => t.kind !== target.kind || t.id !== target.id),
			target,
		];
		this.picker = null;
		this.saveDraft();
		void this.updateComplete.then(() => {
			area.focus({ preventScroll: true });
			area.setSelectionRange(next.caret, next.caret);
		});
	}

	/** Picker keys; true when the key was used. */
	private pickerKey(event: KeyboardEvent): boolean {
		const options = this.pickerOptions();
		if (!this.picker || options.length === 0) return false;
		if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
			const step = event.key === 'ArrowDown' ? 1 : -1;
			this.pickerIndex = (this.pickerIndex + step + options.length) % options.length;
		} else if (
			(event.key === 'Enter' && !event.shiftKey && !event.isComposing) ||
			event.key === 'Tab'
		) {
			this.chooseMention(options[Math.min(this.pickerIndex, options.length - 1)]!);
		} else if (event.key === 'Escape') this.picker = null;
		else return false;
		event.preventDefault();
		return true;
	}

	private onKeydown(event: KeyboardEvent): void {
		if (this.pickerKey(event)) return;
		if (event.key === 'Escape' && this.emojiOpen) {
			this.emojiOpen = false;
			this.focusMessage();
			return;
		}
		if (event.key === 'Enter' && !event.shiftKey && !event.isComposing) {
			event.preventDefault();
			this.submit();
		} else if (event.key === 'Escape' && (this.replyingTo !== null || this.editing)) {
			this.fire('office-chat-cancel', {});
		}
	}

	private pick(event: Event): void {
		const input = event.target as HTMLInputElement;
		this.addFiles(Array.from(input.files ?? []));
		input.value = '';
	}
	private addFiles(files: File[]): void {
		if (this.disabled) return;
		this.files = [...this.files, ...files].slice(0, MAX_FILES);
		this.missingFiles = this.missingFiles.filter(
			(name) => !this.files.some((file) => file.name === name),
		);
		this.saveDraft();
		this.focusMessage();
	}
	private saveDraft(): void {
		this.fire('office-chat-draft', {
			text: this.value,
			files: [...this.files],
			missingFiles: [...this.missingFiles],
		});
	}

	private typingLine(): string {
		const n = this.typing.slice(0, 3);
		if (n.length === 0) return '';
		return n.length === 1 ? `${n[0]} is typing…` : `${n.join(', ')} are typing…`;
	}

	private mentionList() {
		const options = this.pickerOptions();
		if (options.length === 0) return nothing;
		return html`<ul
			id="mention-list"
			class="mentions"
			role="listbox"
			aria-label="People to mention"
			@mousedown=${(event: Event) => event.preventDefault()}
		>
			${options.map(
				(target, index) =>
					html`<li
						id=${`mention-option-${index}`}
						role="option"
						aria-selected=${String(index === this.pickerIndex)}
						@click=${() => this.chooseMention(target)}
					>
						${
							target.kind === 'user'
								? html`<office-ui-avatar
										size="sm"
										name=${target.name}
										seed=${target.id}
									></office-ui-avatar>`
								: html`<span class="everyone" aria-hidden="true">@</span>`
						}
						<span class="name">${target.name}</span>
						${target.kind === 'user' ? nothing : html`<span class="hint">${target.kind === 'channel' ? 'Notify everyone in this channel' : 'Notify the whole team'}</span>`}
					</li>`,
			)}
		</ul>`;
	}

	protected override render() {
		const banner = this.editing
			? 'Editing message'
			: this.replyingTo !== null
				? `Replying to ${this.replyingTo}`
				: '';
		return html`
			<div class="typing" aria-live="polite">${this.typingLine()}</div>
			<div
				class="box"
				@dragover=${(event: DragEvent) => {
					if (event.dataTransfer?.types.includes('Files')) {
						event.preventDefault();
						event.stopPropagation();
					}
				}}
				@drop=${(event: DragEvent) => {
					if (event.dataTransfer?.files.length) {
						event.preventDefault();
						event.stopPropagation();
						this.addFiles([...event.dataTransfer.files]);
					}
				}}
			>
				${
					this.missingFiles.length
						? html`<div role="status">
								Reattach or discard these draft attachments before sending:
								<ul>
									${this.missingFiles.map(
										(name) =>
											html`<li>
												${name}<button
													type="button"
													aria-label=${`Discard missing ${name}`}
													@click=${() => {
														this.missingFiles = this.missingFiles.filter((item) => item !== name);
														this.saveDraft();
													}}
												>
													Discard
												</button>
											</li>`,
									)}
								</ul>
							</div>`
						: nothing
				}
				${
					banner
						? html`<div class="banner">
								<span>${banner}</span>
								<button
									type="button"
									aria-label="Cancel"
									@click=${() => this.fire('office-chat-cancel', {})}
								>
									${icon('close')}
								</button>
							</div>`
						: nothing
				}
				${this.mentionList()}
				<textarea
					rows="1"
					aria-label="Message"
					aria-haspopup="listbox"
					aria-controls=${this.pickerOptions().length ? 'mention-list' : nothing}
					aria-activedescendant=${this.pickerOptions().length ? `mention-option-${this.pickerIndex}` : nothing}
					.value=${live(this.value)}
					placeholder=${this.placeholder}
					?disabled=${this.disabled}
					@input=${this.onInput}
					@paste=${(event: ClipboardEvent) => {
						if (event.clipboardData?.files.length) {
							event.preventDefault();
							this.addFiles([...event.clipboardData.files]);
						}
					}}
					@keydown=${this.onKeydown}
					@keyup=${(event: KeyboardEvent) => {
						if (['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) this.syncPicker();
					}}
					@click=${this.syncPicker}
					@blur=${() => (this.picker = null)}
				></textarea>
				${
					this.files.length
						? html`<ul class="pending">
								${this.files.map(
									(f, i) =>
										html`<li>
											<button
												type="button"
												aria-label=${`Remove ${f.name}`}
												@click=${() => {
													this.files = this.files.filter((_, j) => j !== i);
													this.focusMessage();
													this.saveDraft();
												}}
											>
												<span class="file-label" title=${f.name}>${f.name}</span>${icon('close')}
											</button>
										</li>`,
								)}
							</ul>`
						: nothing
				}
				<div class="bar">
					<button
						type="button"
						title="Attach files"
						aria-label="Attach files"
						?disabled=${this.disabled}
						@click=${() => this.renderRoot.querySelector<HTMLInputElement>('input[type=file]')?.click()}
					>
						${icon('paperclip')}
					</button>
					<div class="emoji">
						<button
							type="button"
							title="Emoji"
							aria-label="Emoji"
							aria-expanded=${String(this.emojiOpen)}
							?disabled=${this.disabled}
							@click=${() => (this.emojiOpen = !this.emojiOpen)}
						>
							${icon('smile')}
						</button>
						${
							this.emojiOpen
								? html`<div class="picker" role="menu">
										${EMOJI.map(
											(e) =>
												html`<button
													type="button"
													role="menuitem"
													@click=${() => {
														const area = this.renderRoot.querySelector('textarea');
														const start = area?.selectionStart ?? this.value.length,
															end = area?.selectionEnd ?? start;
														this.value = this.value.slice(0, start) + e + this.value.slice(end);
														void this.updateComplete.then(() => {
															area?.focus({ preventScroll: true });
															area?.setSelectionRange(start + e.length, start + e.length);
														});
														this.saveDraft();
														this.emojiOpen = false;
														this.renderRoot.querySelector('textarea')?.focus();
													}}
												>
													${e}
												</button>`,
										)}
									</div>`
								: nothing
						}
					</div>
					<span class="spacer"></span>
					<button
						type="button"
						class="send"
						title="Send"
						aria-label="Send message"
						?disabled=${!this.canSend}
						@click=${this.submit}
					>
						${icon('send')}
					</button>
				</div>
				<input type="file" multiple hidden @change=${this.pick} />
			</div>
		`;
	}
}

export const defineChatComposer = definer('office-ui-chat-composer', () => OfficeUiChatComposer);
