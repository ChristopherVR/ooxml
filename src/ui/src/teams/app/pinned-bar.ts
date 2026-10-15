import { LitElement, html, nothing, unsafeCSS } from 'lit';
import type { PinnedMessage } from 'ooxml-core/teams';
import { icon } from '../icons';
import css from './pinned-bar.css?raw';

const SNIPPET = 140;
const snippet = (m: PinnedMessage['message']): string => {
	const text = m.text.replace(/\s+/gu, ' ').trim();
	if (text) return text.length > SNIPPET ? `${text.slice(0, SNIPPET)}...` : text;
	return m.attachments[0]?.name ?? 'Attachment';
};

/**
 * The pinned bar at the top of a channel, like Teams: the newest pin on one line and a count that
 * expands the full list. Properties: `pinned` (newest first). Events (bubbling, composed):
 * `teams-pinned-open` `{ messageId }` (jump to the message) and `teams-pinned-unpin` `{ messageId }`.
 */
export class TeamsPinnedBar extends LitElement {
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
			<section aria-label="Pinned messages">
				<div class="summary">
					<button
						type="button"
						class="latest"
						@click=${() => this.emit('teams-pinned-open', latest.message.id)}
					>
						<span class="mark">${icon('pin')}</span>
						<strong>${latest.message.authorName}</strong>
						<span class="text">${snippet(latest.message)}</span>
					</button>
					<button
						type="button"
						class="count"
						aria-expanded=${String(this.expanded)}
						aria-controls="pins"
						@click=${() => (this.expanded = !this.expanded)}
					>
						${this.pinned.length} pinned
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
												<span class="text">${snippet(pin.message)}</span>
												<span class="by">Pinned by ${pin.pinnedByName}</span>
											</button>
											<button
												type="button"
												class="unpin"
												aria-label=${`Unpin message from ${pin.message.authorName}`}
												title="Unpin"
												@click=${() => this.emit('teams-pinned-unpin', pin.message.id)}
											>
												Unpin
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
