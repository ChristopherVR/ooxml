import { html, nothing } from 'lit';
import { repeat } from 'lit/directives/repeat.js';
import { definer } from '../registry';
import { TeamsElement, withStyles } from './base';
import { icon } from './icons';
import css from './call-grid.css?raw';

export interface CallTile {
	id: string;
	name: string;
	self?: boolean;
	stream?: MediaStream | null;
	audio?: boolean;
	video?: boolean;
	screen?: boolean;
	hand?: boolean;
	/** `connecting` and `failed` show a status chip; anything else is treated as connected. */
	connection?: string;
}

const gridColumns = (n: number): number => Math.max(1, Math.ceil(Math.sqrt(n)));

/**
 * The meeting stage. Property `participants` (`CallTile[]`). Someone sharing their screen is
 * spotlighted with everyone else in a side strip. Your own tile is muted (no echo) and mirrored
 * unless it shows a screen; tiles without video show the avatar. The grid only attaches the
 * streams it is handed; capture and connections belong to the host.
 */
export class OfficeUiCallGrid extends TeamsElement {
	static override styles = withStyles(css);
	static override properties = { participants: { attribute: false } };
	declare participants: CallTile[];

	constructor() {
		super();
		this.participants = [];
	}

	override connectedCallback(): void {
		super.connectedCallback();
		this.setAttribute('role', 'list');
		this.setAttribute('aria-label', 'Participants');
	}

	protected override render() {
		const people = this.participants;
		if (people.length === 0) return html`<p class="empty">Nobody is in the call</p>`;
		const sharer = people.find((p) => p.screen);
		const spotlight = sharer !== undefined && people.length > 1;
		const ordered = sharer ? [sharer, ...people.filter((p) => p !== sharer)] : people;
		const style = `--cols:${gridColumns(people.length)};--rows:${Math.max(1, people.length - 1)}`;
		return html`
			<div class="grid" data-spotlight=${String(spotlight)} style=${style}>
				${repeat(
					ordered,
					(p) => p.id,
					(p, i) => this.tile(p, spotlight && i === 0),
				)}
			</div>
		`;
	}

	private tile(p: CallTile, main: boolean) {
		const showVideo = Boolean(p.stream) && (p.video || p.screen || !p.self);
		const showAvatar = !showVideo || (!p.video && !p.screen);
		return html`
			<div
				class="tile"
				role="listitem"
				data-self=${String(Boolean(p.self))}
				data-screen=${String(Boolean(p.screen))}
				data-main=${String(main)}
			>
				${
					showVideo
						? html`<video
								autoplay
								playsinline
								.muted=${Boolean(p.self)}
								.srcObject=${p.stream ?? null}
							></video>`
						: nothing
				}
				${showAvatar ? html`<office-ui-avatar size="xl" name=${p.name} seed=${p.id}></office-ui-avatar>` : nothing}
				<div class="name">
					${p.audio === false ? html`<span class="muted">${icon('micOff')}</span>` : nothing}
					<span>${p.self ? `${p.name} (you)` : p.name}</span>
				</div>
				${p.hand ? html`<span class="hand" role="img" aria-label="Hand raised">✋</span>` : nothing}
				${
					!p.self && p.connection && p.connection !== 'connected'
						? html`<span class="status"
								>${p.connection === 'failed' ? 'Connection failed' : 'Connecting…'}</span
							>`
						: nothing
				}
			</div>
		`;
	}
}

export const defineCallGrid = definer('office-ui-call-grid', () => OfficeUiCallGrid);
