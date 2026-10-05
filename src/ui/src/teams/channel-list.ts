import { html } from 'lit';
import { definer } from '../registry.js';
import { TeamsElement, withStyles } from './base.js';
import { icon } from './icons.js';
import css from './channel-list.css?raw';

export interface ChannelRow {
	id: string;
	name: string;
	topic?: string;
	kind?: 'channel' | 'direct';
	unread?: number;
	/** Someone is in a call in this channel. */
	live?: boolean;
}

/**
 * The team and its channels, like the Teams navigation tree. Properties: `heading` (the team
 * name), `channels`, `selectedId`. Emits `office-channel-select` `{ id }` and
 * `office-channel-create` `{}` (the host asks for a name). Unread channels are bold with a count.
 */
export class OfficeUiChannelList extends TeamsElement {
	static override styles = withStyles(css);
	static override properties = {
		heading: { type: String },
		channels: { attribute: false },
		selectedId: { type: String, attribute: 'selected-id' },
		collapsed: { state: true },
	};
	declare heading: string;
	declare channels: ChannelRow[];
	declare selectedId: string;
	declare collapsed: boolean;

	constructor() {
		super();
		this.heading = 'Channels';
		this.channels = [];
		this.selectedId = '';
		this.collapsed = false;
	}

	protected override render() {
		return html`
			<div class="team">
				<button
					class="toggle"
					type="button"
					aria-expanded=${String(!this.collapsed)}
					@click=${() => (this.collapsed = !this.collapsed)}
				>
					<span class="chevron" data-open=${String(!this.collapsed)}>${icon('chevronRight')}</span>
					<span class="name">${this.heading}</span>
				</button>
				<button
					class="add"
					type="button"
					title="Add channel"
					aria-label="Add channel"
					@click=${() => this.fire('office-channel-create', {})}
				>
					${icon('plus')}
				</button>
			</div>
			<ul ?hidden=${this.collapsed}>
				${this.channels.map(
					(c) => html`
						<li>
							<button
								class="row"
								type="button"
								aria-current=${c.id === this.selectedId ? 'true' : 'false'}
								data-unread=${c.unread ? 'true' : 'false'}
								@click=${() => this.fire('office-channel-select', { id: c.id })}
							>
								${icon(c.kind === 'direct' ? 'message' : 'hash')}
								<span class="title">${c.name}</span>
								${
									c.live
										? html`<span class="live" title="Call in progress">${icon('video')}</span>`
										: ''
								}
								${
									c.unread
										? html`<span class="count">${c.unread > 99 ? '99+' : c.unread}</span>`
										: ''
								}
							</button>
						</li>
					`,
				)}
			</ul>
		`;
	}
}

export const defineChannelList = definer('office-ui-channel-list', () => OfficeUiChannelList);
