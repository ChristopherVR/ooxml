import { html, nothing, type PropertyValues } from 'lit';
import { definer } from '../registry.js';
import { TeamsElement, withStyles } from './base.js';
import { icon } from './icons.js';
import css from './call-controls.css?raw';

export type CallAction = 'chat' | 'people' | 'hand' | 'camera' | 'mic' | 'screen' | 'leave';

export const formatElapsed = (ms: number): string => {
	const total = Math.max(0, Math.floor(ms / 1000));
	const pad = (n: number): string => String(n).padStart(2, '0');
	const h = Math.floor(total / 3600);
	return `${h ? `${h}:` : ''}${pad(Math.floor((total % 3600) / 60))}:${pad(total % 60)}`;
};

/**
 * The meeting bar: elapsed time on the left, the controls in the middle (Chat, People, Raise
 * hand, Camera, Mic, Share, Leave). Properties: `title`, `startedAt` (epoch ms), `mic`, `camera`,
 * `screen`, `hand`, `panel` (`'chat' | 'people' | ''`), `shareUnavailable`. Emits
 * `office-call-action` `{ action }`.
 */
export class OfficeUiCallControls extends TeamsElement {
	static override styles = withStyles(css);
	static override properties = {
		heading: { type: String },
		startedAt: { type: Number, attribute: 'started-at' },
		mic: { type: Boolean },
		camera: { type: Boolean },
		screen: { type: Boolean },
		hand: { type: Boolean },
		panel: { type: String },
		shareUnavailable: { type: Boolean, attribute: 'share-unavailable' },
		now: { state: true },
	};
	declare heading: string;
	declare startedAt: number;
	declare mic: boolean;
	declare camera: boolean;
	declare screen: boolean;
	declare hand: boolean;
	declare panel: string;
	declare shareUnavailable: boolean;
	declare now: number;
	private timer: ReturnType<typeof setInterval> | undefined;

	constructor() {
		super();
		this.heading = '';
		this.startedAt = 0;
		this.mic = false;
		this.camera = false;
		this.screen = false;
		this.hand = false;
		this.panel = '';
		this.shareUnavailable = false;
		this.now = Date.now();
	}

	override connectedCallback(): void {
		super.connectedCallback();
		this.setAttribute('role', 'toolbar');
		this.setAttribute('aria-label', 'Meeting controls');
		this.timer = setInterval(() => (this.now = Date.now()), 1000);
	}
	override disconnectedCallback(): void {
		clearInterval(this.timer);
		super.disconnectedCallback();
	}

	protected override willUpdate(_changed: PropertyValues<this>): void {
		// `now` only ticks while a start time is known.
		if (!this.startedAt) this.now = this.startedAt;
	}

	private button(
		action: CallAction,
		label: string,
		glyph: string,
		state: { on?: boolean; off?: boolean; disabled?: boolean } = {},
	) {
		return html`
			<button
				type="button"
				data-action=${action}
				data-off=${String(Boolean(state.off))}
				aria-label=${label}
				aria-pressed=${state.on === undefined ? nothing : String(state.on)}
				title=${label}
				?disabled=${state.disabled}
				@click=${() => this.fire('office-call-action', { action })}
			>
				${icon(glyph)}<span class="text">${label}</span>
			</button>
		`;
	}

	protected override render() {
		return html`
			<div class="info">
				<span class="heading">${this.heading}</span
				>${this.startedAt ? html`<time>${formatElapsed(this.now - this.startedAt)}</time>` : nothing}
			</div>
			<div class="actions">
				${this.button('chat', 'Chat', 'chat', { on: this.panel === 'chat' })}
				${this.button('people', 'People', 'users', { on: this.panel === 'people' })}
				${this.button('hand', this.hand ? 'Lower' : 'Raise', 'hand', { on: this.hand })}
				${this.button('camera', this.camera ? 'Camera off' : 'Camera on', this.camera ? 'video' : 'videoOff', { off: !this.camera })}
				${this.button('mic', this.mic ? 'Mute' : 'Unmute', this.mic ? 'mic' : 'micOff', { off: !this.mic })}
				${this.button('screen', this.screen ? 'Stop sharing' : 'Share', 'screenShare', { on: this.screen, disabled: this.shareUnavailable })}
				<button
					type="button"
					class="leave"
					aria-label="Leave call"
					@click=${() => this.fire('office-call-action', { action: 'leave' })}
				>
					${icon('phoneOff')}<span class="text">Leave</span>
				</button>
			</div>
		`;
	}
}

export const defineCallControls = definer('office-ui-call-controls', () => OfficeUiCallControls);
