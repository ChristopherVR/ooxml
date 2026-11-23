import { html } from 'lit';
import { OfficeElement, controlStyles } from '../../base';
import { defineElement } from '../../registry';
import css from './cameo-preview.css?raw';
import { RIBBON_ICON_PATHS } from './ribbon-icons';

export class PptxUiCameoPreview extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		active: { state: true },
		pending: { state: true },
		unavailable: { state: true },
		cameraLabel: { type: String, attribute: 'camera-label' },
		unavailableLabel: { type: String, attribute: 'unavailable-label' },
	};

	declare cameraLabel?: string;
	declare unavailableLabel?: string;
	declare private active: boolean;
	declare private pending: boolean;
	declare private unavailable: boolean;
	#request = 0;
	#stream: MediaStream | undefined;

	override disconnectedCallback(): void {
		this.#stop();
		super.disconnectedCallback();
	}

	override render() {
		return html`
			<video ?hidden=${!this.active} muted playsinline autoplay aria-hidden="true"></video>
			<button
				type="button"
				aria-label=${this.cameraLabel ?? 'Cameo'}
				aria-pressed=${Boolean(this.active)}
				?disabled=${this.pending}
				@pointerdown=${(event: Event) => event.stopPropagation()}
				@keydown=${this.#keyDown}
				@click=${this.#toggle}
			>
				<svg class="camera-icon" viewBox="0 0 20 20" aria-hidden="true">
					<path d=${RIBBON_ICON_PATHS.camera}></path>
				</svg>
			</button>
			${
				this.unavailable
					? html`<span role="status">${this.unavailableLabel ?? 'Camera unavailable'}</span>`
					: ''
			}
		`;
	}

	#keyDown = (event: KeyboardEvent) => {
		if (event.key === ' ' || event.key === 'Enter') event.stopPropagation();
	};

	#toggle = async (event: Event) => {
		event.stopPropagation();
		if (this.active) {
			this.#stop();
			return;
		}
		if (this.pending) return;
		this.unavailable = false;
		this.pending = true;
		const request = ++this.#request;
		try {
			const stream = await this.ownerDocument.defaultView?.navigator.mediaDevices?.getUserMedia({
				video: true,
				audio: false,
			});
			if (!stream) throw new Error('Camera unavailable');
			const video = this.renderRoot.querySelector('video');
			if (request !== this.#request || !this.isConnected || !video) {
				for (const track of stream.getTracks()) track.stop();
				return;
			}
			this.#stream = stream;
			video.srcObject = stream;
			this.active = true;
		} catch {
			if (request === this.#request) {
				this.#stop();
				this.unavailable = true;
			}
		} finally {
			if (request === this.#request) this.pending = false;
		}
	};

	#stop() {
		this.#request++;
		for (const track of this.#stream?.getTracks() ?? []) track.stop();
		this.#stream = undefined;
		const video = this.renderRoot?.querySelector('video');
		if (video) video.srcObject = null;
		this.active = false;
		this.pending = false;
	}
}

export function definePptxCameoPreview(registry: CustomElementRegistry): void {
	defineElement(registry, 'pptx-ui-cameo-preview', PptxUiCameoPreview);
}

declare global {
	interface HTMLElementTagNameMap {
		'pptx-ui-cameo-preview': PptxUiCameoPreview;
	}
}
