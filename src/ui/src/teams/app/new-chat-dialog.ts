import { html, nothing, type PropertyValues } from 'lit';
import { type ContactView, MAX_CHANNEL_NAME_CHARS, MAX_CHAT_MEMBERS } from 'ooxml-core/teams';
import { dialogStyles, TeamsDialog } from './dialog-base';
import css from './new-chat-dialog.css?raw';

/**
 * New chat: pick one person for a 1:1 chat, or several for a group chat with an optional name.
 * Properties: `open`, `contacts`, `create(userIds, name)` (true when the chat was made).
 * Emits `teams-dialog-close`.
 */
export class TeamsNewChatDialog extends TeamsDialog {
	static override styles = dialogStyles(css);
	static override properties = {
		...TeamsDialog.properties,
		contacts: { attribute: false },
		create: { attribute: false },
		selected: { state: true },
		name: { state: true },
		issue: { state: true },
	};
	declare contacts: ContactView[];
	declare create: ((userIds: string[], name: string) => boolean) | undefined;
	declare selected: string[];
	declare name: string;
	declare issue: string;

	constructor() {
		super();
		this.contacts = [];
		this.create = undefined;
		this.selected = [];
		this.name = '';
		this.issue = '';
	}

	protected override willUpdate(changed: PropertyValues<this>): void {
		if (changed.has('open') && this.open) {
			this.selected = [];
			this.name = '';
			this.issue = '';
		}
	}

	private submit(event: Event): void {
		event.preventDefault();
		if (this.selected.length === 0) {
			this.issue = this.t('newChat.pickOne');
			return;
		}
		if (!this.create?.(this.selected, this.selected.length > 1 ? this.name : ''))
			this.issue = this.t('newChat.failed');
	}

	protected override render() {
		const group = this.selected.length > 1;
		return this.frame(
			this.t('chat.new'),
			html`<form @submit=${this.submit}>
				<div class="field">
					${this.t('newChat.to')}
					<office-ui-people-picker
						label=${this.t('newChat.peopleLabel')}
						.contacts=${this.contacts}
						.selected=${this.selected}
						@office-people-change=${(e: CustomEvent<{ selected: string[] }>) => {
							this.selected = e.detail.selected.slice(0, MAX_CHAT_MEMBERS - 1);
							this.issue = '';
						}}
					></office-ui-people-picker>
				</div>
				${
					group
						? html`<label
								>${this.t('chatPeople.groupNameOptional')}<input
									maxlength=${MAX_CHANNEL_NAME_CHARS}
									.value=${this.name}
									placeholder=${this.t('chatPeople.groupPlaceholder')}
									@input=${(e: Event) => (this.name = (e.target as HTMLInputElement).value)}
							/></label>`
						: nothing
				}
				<small>${group ? this.t('newChat.hintGroup') : this.t('newChat.hintOne')}</small>
				${this.issue ? html`<p class="issue" role="alert">${this.issue}</p>` : nothing}
				<footer>
					<button type="button" @click=${this.requestClose}>${this.t('common.cancel')}</button>
					<button type="submit" class="primary" ?disabled=${this.selected.length === 0}>
						${this.t(group ? 'newChat.startGroup' : 'newChat.start')}
					</button>
				</footer>
			</form>`,
			this.t('dialog.closeNewChat'),
		);
	}
}

export function defineTeamsNewChatDialog(
	registry: CustomElementRegistry | undefined = globalThis.customElements,
): void {
	if (registry && !registry.get('teams-new-chat-dialog'))
		registry.define('teams-new-chat-dialog', TeamsNewChatDialog);
}
