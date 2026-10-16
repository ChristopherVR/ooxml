import { html, nothing, type PropertyValues } from 'lit';
import { type ChatView, type ContactView, MAX_CHANNEL_NAME_CHARS } from 'ooxml-core/teams';
import { dialogStyles, TeamsDialog } from './dialog-base';
import css from './chat-people-dialog.css?raw';

/**
 * The people of a chat: who is in it, rename (groups), add and remove people, leave (groups).
 * Properties: `open`, `chat`, `contacts`. Emits `teams-chat-rename` `{ id, name }`,
 * `teams-chat-add` `{ id, userIds }`, `teams-chat-remove` `{ id, userId }`,
 * `teams-chat-leave` `{ id }` and `teams-dialog-close`.
 */
export class TeamsChatPeopleDialog extends TeamsDialog {
	static override styles = dialogStyles(css);
	static override properties = {
		...TeamsDialog.properties,
		chat: { attribute: false },
		contacts: { attribute: false },
		name: { state: true },
		adding: { state: true },
	};
	declare chat: ChatView | null;
	declare contacts: ContactView[];
	declare name: string;
	declare adding: string[];

	constructor() {
		super();
		this.chat = null;
		this.contacts = [];
		this.name = '';
		this.adding = [];
	}

	protected override willUpdate(changed: PropertyValues<this>): void {
		if (changed.has('open') && this.open) {
			this.name = this.chat?.name ?? '';
			this.adding = [];
		}
	}

	protected override render() {
		const chat = this.chat;
		if (!chat) return this.frame(this.t('chat.people'), html``, this.t('dialog.closePeople'));
		const group = chat.kind === 'group';
		return this.frame(
			this.t('chat.people'),
			html`<div class="body">
				${
					group
						? html`<form
								class="rename"
								@submit=${(e: Event) => {
									e.preventDefault();
									this.fire('teams-chat-rename', { id: chat.id, name: this.name });
								}}
							>
								<label
									>${this.t('chatPeople.groupName')}<input
										maxlength=${MAX_CHANNEL_NAME_CHARS}
										.value=${this.name}
										placeholder=${this.t('chatPeople.groupPlaceholder')}
										@input=${(e: Event) => (this.name = (e.target as HTMLInputElement).value)}
								/></label>
								<button type="submit">${this.t('chatPeople.rename')}</button>
							</form>`
						: nothing
				}
				<h3>${this.t('chatPeople.inChat', { count: chat.members.length })}</h3>
				<ul class="members">
					${chat.members.map(
						(m) =>
							html`<li>
								<office-ui-avatar
									size="sm"
									name=${m.name}
									seed=${m.id}
									presence=${m.availability}
								></office-ui-avatar>
								<span class="name"
									>${m.self ? this.t('common.nameYou', { name: m.name }) : m.name}</span
								>
								${
									group && !m.self && chat.members.length > 1
										? html`<button
												type="button"
												aria-label=${this.t('people.remove', { name: m.name })}
												@click=${() => this.fire('teams-chat-remove', { id: chat.id, userId: m.id })}
											>
												${this.t('chatPeople.remove')}
											</button>`
										: nothing
								}
							</li>`,
					)}
				</ul>
				${
					group
						? html`<h3>${this.t('chatPeople.addPeople')}</h3>
								<office-ui-people-picker
									label=${this.t('chatPeople.peopleToAdd')}
									.contacts=${this.contacts}
									.exclude=${chat.members.map((m) => m.id)}
									.selected=${this.adding}
									@office-people-change=${(e: CustomEvent<{ selected: string[] }>) => (this.adding = e.detail.selected)}
								></office-ui-people-picker>
								<footer>
									<button
										type="button"
										class="primary"
										?disabled=${this.adding.length === 0}
										@click=${() => {
											this.fire('teams-chat-add', { id: chat.id, userIds: this.adding });
											this.adding = [];
										}}
									>
										${this.t('chatPeople.addToChat')}
									</button>
								</footer>`
						: html`<p>${this.t('chatPeople.directNote')}</p>`
				}
				<footer>
					${
						chat.canLeave
							? html`<button
									type="button"
									class="danger"
									@click=${() => this.fire('teams-chat-leave', { id: chat.id })}
								>
									${this.t('chatPeople.leave')}
								</button>`
							: nothing
					}
					<button type="button" @click=${this.requestClose}>${this.t('common.done')}</button>
				</footer>
			</div>`,
			this.t('dialog.closePeople'),
		);
	}
}

export function defineTeamsChatPeopleDialog(
	registry: CustomElementRegistry | undefined = globalThis.customElements,
): void {
	if (registry && !registry.get('teams-chat-people-dialog'))
		registry.define('teams-chat-people-dialog', TeamsChatPeopleDialog);
}
