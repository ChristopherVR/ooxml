import { html, type PropertyValues } from 'lit';
import { definer } from '../registry';
import { initialsOf } from '../presence';
import { TeamsElement, colorFor, withStyles } from './base';
import css from './avatar.css?raw';

export type Presence = 'available' | 'busy' | 'dnd' | 'brb' | 'away' | 'offline';
const PRESENCE_KEY = {
	available: 'presence.available',
	busy: 'presence.busy',
	dnd: 'presence.dnd',
	brb: 'presence.brb',
	away: 'presence.away',
	offline: 'presence.offline',
} as const satisfies Record<Presence, string>;

/**
 * Round avatar with initials and an optional presence badge, like the Teams persona.
 * Attributes: `name`, `color` (default derived from `seed`/`name`), `size` (`sm` | `md` | `lg` |
 * `xl`), `presence` (`available` | `busy` | `dnd` | `brb` | `away` | `offline`, absent = no badge).
 */
export class OfficeUiAvatar extends TeamsElement {
	static override styles = withStyles(css);
	static override properties = {
		name: { type: String },
		seed: { type: String },
		color: { type: String },
		size: { type: String, reflect: true },
		presence: { type: String, reflect: true },
	};
	declare name: string;
	declare seed: string;
	declare color: string;
	declare size: 'sm' | 'md' | 'lg' | 'xl';
	declare presence: Presence | '';

	constructor() {
		super();
		this.name = '';
		this.seed = '';
		this.color = '';
		this.size = 'md';
		this.presence = '';
	}

	protected override willUpdate(changed: PropertyValues<this>): void {
		if (changed.has('name') || changed.has('presence') || changed.has('locale')) {
			this.setAttribute('role', 'img');
			const status = this.presence ? `, ${this.t(PRESENCE_KEY[this.presence])}` : '';
			this.setAttribute('aria-label', `${this.name}${status}`);
		}
	}

	protected override render() {
		const background = this.color || colorFor(this.seed || this.name);
		return html`
			<span class="face" style="background:${background}">${initialsOf(this.name)}</span>
			${
				this.presence
					? html`<span
							class="presence"
							data-presence=${this.presence}
							title=${this.t(PRESENCE_KEY[this.presence])}
						></span>`
					: ''
			}
		`;
	}
}

export const defineAvatar = definer('office-ui-avatar', () => OfficeUiAvatar);
