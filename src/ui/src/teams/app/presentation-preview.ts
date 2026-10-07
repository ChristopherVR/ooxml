import { LitElement, html, nothing, unsafeCSS } from 'lit';
import type { PptxData, PptxHandler, SvgExporter } from 'ooxml-core/pptx';
import css from './presentation-preview.css?raw';

/** Static slide previews reuse the core exporter and render its SVG as an isolated image. */
export class TeamsPresentationPreview extends LitElement {
	static override styles = unsafeCSS(css);
	static override properties = {
		data: { state: true },
		slide: { state: true },
		image: { state: true },
		warnings: { state: true },
	};
	declare data: PptxData | null;
	declare slide: number;
	declare image: string;
	declare warnings: string[];
	private handler: PptxHandler | undefined;
	private exporter: typeof SvgExporter | undefined;
	private generation = 0;
	private readonly metrics = document.createElement('canvas').getContext('2d');

	constructor() {
		super();
		this.data = null;
		this.slide = 0;
		this.image = '';
		this.warnings = [];
	}
	async load(bytes: Uint8Array): Promise<void> {
		const generation = ++this.generation;
		this.clear();
		this.data = null;
		const { PptxHandler, SvgExporter } = await import('ooxml-core/pptx');
		if (generation !== this.generation || !this.isConnected) return;
		const handler = new PptxHandler();
		try {
			const data = await handler.load(new Uint8Array(bytes).buffer, {
				eagerDecodeImages: true,
				allowExternalImages: false,
				maxUncompressedBytes: 134_217_728,
			});
			if (generation !== this.generation || !this.isConnected) {
				handler.dispose();
				return;
			}
			if (!data.slides.length) throw new Error('The presentation has no slides');
			this.handler = handler;
			this.exporter = SvgExporter;
			this.slide = 0;
			this.data = data;
			this.warnings = [
				...new Set(
					[...(data.warnings ?? []), ...data.slides.flatMap((slide) => slide.warnings ?? [])].map(
						(warning) => warning.message,
					),
				),
			];
			this.showSlide();
			await this.updateComplete;
		} catch (error) {
			if (this.handler === handler) {
				this.clear();
				this.data = null;
			} else handler.dispose();
			throw error;
		}
	}
	private clear(): void {
		if (this.image) URL.revokeObjectURL(this.image);
		this.image = '';
		this.handler?.dispose();
		this.handler = undefined;
		this.exporter = undefined;
	}
	override disconnectedCallback(): void {
		this.generation++;
		this.clear();
		super.disconnectedCallback();
	}
	private showSlide(): void {
		const data = this.data,
			slide = data?.slides[this.slide];
		if (!data || !slide || !this.exporter) return;
		const svg = this.exporter.exportSlide(slide, data.width, data.height, {
			measureText: (text, font) => {
				if (!this.metrics) return Number.NaN;
				this.metrics.font = `${font.italic ? 'italic ' : ''}${font.bold ? 'bold ' : ''}${font.size}px ${JSON.stringify(font.family)}`;
				return this.metrics.measureText(text).width;
			},
		});
		const image = URL.createObjectURL(new Blob([svg], { type: 'image/svg+xml' }));
		if (this.image) URL.revokeObjectURL(this.image);
		this.image = image;
	}
	private move(delta: number): void {
		this.slide = Math.max(0, Math.min((this.data?.slides.length ?? 1) - 1, this.slide + delta));
		this.showSlide();
	}
	protected override render() {
		const data = this.data;
		if (!data) return nothing;
		const slide = data.slides[this.slide]!;
		return html`<p class="hint">
				Static slide preview. Animation, transitions, media playback and editing are unavailable.
				Layout and fonts may differ from PowerPoint.
			</p>
			<nav aria-label="Slide navigation">
				<button type="button" ?disabled=${this.slide === 0} @click=${() => this.move(-1)}>
					Previous slide
				</button>
				<span role="status"
					>Slide ${this.slide + 1} of ${data.slides.length}${slide.hidden ? ' (hidden)' : ''}</span
				>
				<button
					type="button"
					?disabled=${this.slide === data.slides.length - 1}
					@click=${() => this.move(1)}
				>
					Next slide
				</button>
			</nav>
			<div class="stage">
				<img
					src=${this.image}
					alt=${`Slide ${this.slide + 1}`}
					@error=${() => (this.warnings = [...this.warnings, 'The slide image could not be displayed.'])}
				/>
			</div>
			<details>
				<summary>Slide text</summary>
				<pre>
${slide.elements
						.filter((element) => 'text' in element && element.text)
						.map((element) => ('text' in element ? element.text : ''))
						.join('\n\n')}</pre>
			</details>
			${
				slide.notes
					? html`<details>
							<summary>Speaker notes</summary>
							<pre>${slide.notes}</pre>
						</details>`
					: nothing
			}
			${
				this.warnings.length
					? html`<details open>
							<summary>Compatibility warnings</summary>
							<ul>
								${this.warnings.map((warning) => html`<li>${warning}</li>`)}
							</ul>
						</details>`
					: nothing
			}`;
	}
}

export function defineTeamsPresentationPreview(): void {
	if (globalThis.customElements && !customElements.get('teams-presentation-preview'))
		customElements.define('teams-presentation-preview', TeamsPresentationPreview);
}
