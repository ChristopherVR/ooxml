import { html, nothing } from 'lit';
import { definer } from '../registry.js';
import { TeamsElement, withStyles } from './base.js';
import { icon } from './icons.js';
import css from './prejoin.css?raw';

/**
 * The pre-join screen: a camera preview and the audio/video switches, then "Join now".
 * Properties: `heading`, `name` (you), `preview` (a `MediaStream` or null), `audio`, `video`,
 * `busy` (joining). Events: `office-prejoin-change` `{ audio?, video? }`, `office-prejoin-join`
 * and `office-prejoin-cancel`.
 */
export class OfficeUiPrejoin extends TeamsElement {
	static override styles = withStyles(css);
	static override properties = {
		heading: { type: String },
		name: { type: String },
		preview: { attribute: false },
		audio: { type: Boolean },
		video: { type: Boolean },
		busy: { type: Boolean },
	};
	declare heading: string;
	declare name: string;
	declare preview: MediaStream | null;
	declare audio: boolean;
	declare video: boolean;
	declare busy: boolean;

	constructor() {
		super();
		this.heading = '';
		this.name = '';
		this.preview = null;
		this.audio = true;
		this.video = true;
		this.busy = false;
	}

	private toggle(label: string, glyph: string, on: boolean, key: 'audio' | 'video') {
		return html`
			<label class="switch">
				<span class="what">${icon(glyph)} ${label}</span>
				<input
					type="checkbox"
					role="switch"
					.checked=${on}
					aria-label=${label}
					@change=${(e: Event) =>
						this.fire('office-prejoin-change', { [key]: (e.target as HTMLInputElement).checked })}
				/>
				<span class="track" aria-hidden="true"></span>
			</label>
		`;
	}

	protected override render() {
		const showVideo = this.video && this.preview;
		return html`
			<section class="card">
				<div class="stage">
					${
						showVideo
							? html`<video
									autoplay
									playsinline
									muted
									.muted=${true}
									.srcObject=${this.preview}
								></video>`
							: html`<office-ui-avatar size="xl" name=${this.name}></office-ui-avatar>
									<p>${this.video ? 'Starting camera…' : 'Your camera is turned off'}</p>`
					}
				</div>
				<div class="panel">
					<h2>Choose your audio and video settings</h2>
					<p class="for">${this.heading}</p>
					${this.toggle('Camera', this.video ? 'video' : 'videoOff', this.video, 'video')}
					${this.toggle('Microphone', this.audio ? 'mic' : 'micOff', this.audio, 'audio')}
					<div class="actions">
						<button
							type="button"
							class="cancel"
							@click=${() => this.fire('office-prejoin-cancel', {})}
						>
							Cancel
						</button>
						<button
							type="button"
							class="join"
							?disabled=${this.busy}
							@click=${() => this.fire('office-prejoin-join', {})}
						>
							${this.busy ? 'Joining…' : 'Join now'}
						</button>
					</div>
					${nothing}
				</div>
			</section>
		`;
	}
}

export const definePrejoin = definer('office-ui-prejoin', () => OfficeUiPrejoin);
