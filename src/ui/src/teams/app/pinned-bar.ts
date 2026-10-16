import { html, nothing, unsafeCSS } from 'lit';
import { TeamsLitElement } from '../i18n/lit-element';
import type { PinnedMessage } from 'ooxml-core/teams';
import { icon } from '../icons';
import css from './pinned-bar.css?raw';

const SNIPPET = 140;
const snippet = (m: PinnedMessage['message'], attachment: string): string => {
	const text = m.text.replace(/\s+/gu, ' ').trim();
	if (text) return text.length > SNIPPET ? `${text.slice(0, SNIPPET)}...` : text;
	return m.attachments[0]?.name ?? attachment;
};

/**
 * The pinned bar at the top of a channel, like Teams: the newest pin on one line and a count that
 * expands the full list. Properties: `pinned` (newest first). Events (bubbling, composed):
 * `teams-pinned-open` `{ messageId }` (jump to the message) and `teams-pinned-unpin` `{ messageId }`.
 */
export class TeamsPinnedBar extends TeamsLitElement {
	static override styles = unsafeCSS(css);
	static override properties = { pinned: { attribute: false }, expanded: { state: true } };
	declare pinned: PinnedMessage[];
	declare expanded: boolean;

	constructor() {
		super();
		this.pinned = [];
		this.expanded = false;
	}

	private emit(name: string, messageId: string): void {
		this.dispatchEvent(
			new CustomEvent(name, { detail: { messageId }, bubbles: true, composed: true }),
		);
	}

	protected override render() {
		const [latest] = this.pinned;
		if (!latest) return nothing;
		return html`
			<section aria-label=${this.t('pinned.messages')}>
				<div class="summary">
					<button
						type="button"
						class="latest"
						@click=${() => this.emit('teams-pinned-open', latest.message.id)}
					>
						<span class="mark">${icon('pin')}</span>
						<strong>${latest.message.authorName}</strong>
						<span class="text">${snippet(latest.message, this.t('pinned.attachment'))}</span>
					</button>
					<button
						type="button"
						class="count"
						aria-expanded=${String(this.expanded)}
						aria-controls="pins"
						@click=${() => (this.expanded = !this.expanded)}
					>
						${this.t('pinned.count', { count: this.pinned.length })}
					</button>
				</div>
				${
					this.expanded
						? html`<ul id="pins">
								${this.pinned.map(
									(pin) =>
										html`<li>
											<button
												type="button"
												class="jump"
												@click=${() => this.emit('teams-pinned-open', pin.message.id)}
											>
												<strong>${pin.message.authorName}</strong>
												<span class="text"
													>${snippet(pin.message, this.t('pinned.attachment'))}</span
												>
												<span class="by">${this.t('pinned.by', { name: pin.pinnedByName })}</span>
											</button>
											<button
												type="button"
												class="unpin"
												aria-label=${this.t('pinned.unpinFrom', { name: pin.message.authorName })}
												title=${this.t('chat.unpin')}
												@click=${() => this.emit('teams-pinned-unpin', pin.message.id)}
											>
												${this.t('chat.unpin')}
											</button>
										</li>`,
								)}
							</ul>`
						: nothing
				}
			</section>
		`;
	}
}

export function defineTeamsPinnedBar(): void {
	if (!customElements.get('teams-pinned-bar'))
		customElements.define('teams-pinned-bar', TeamsPinnedBar);
}
