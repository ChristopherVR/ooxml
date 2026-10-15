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
			this.issue = 'Pick at least one person.';
			return;
		}
		if (!this.create?.(this.selected, this.selected.length > 1 ? this.name : ''))
			this.issue = 'The chat could not be started.';
	}

	protected override render() {
		const group = this.selected.length > 1;
		return this.frame(
			'New chat',
			html`<form @submit=${this.submit}>
				<div class="field">
					To
					<office-ui-people-picker
						label="People to chat with"
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
								>Group name (optional)<input
									maxlength=${MAX_CHANNEL_NAME_CHARS}
									.value=${this.name}
									placeholder="Name this group chat"
									@input=${(e: Event) => (this.name = (e.target as HTMLInputElement).value)}
							/></label>`
						: nothing
				}
				<small
					>${
						group
							? 'Everyone you pick joins one group chat. Only its members can read it.'
							: 'Pick one person for a 1:1 chat, or several for a group chat.'
					}</small
				>
				${this.issue ? html`<p class="issue" role="alert">${this.issue}</p>` : nothing}
				<footer>
					<button type="button" @click=${this.requestClose}>Cancel</button>
					<button type="submit" class="primary" ?disabled=${this.selected.length === 0}>
						${group ? 'Start group chat' : 'Start chat'}
					</button>
				</footer>
			</form>`,
		);
	}
}

export function defineTeamsNewChatDialog(
	registry: CustomElementRegistry | undefined = globalThis.customElements,
): void {
	if (registry && !registry.get('teams-new-chat-dialog'))
		registry.define('teams-new-chat-dialog', TeamsNewChatDialog);
}
