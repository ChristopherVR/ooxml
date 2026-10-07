import { LitElement, html, nothing, unsafeCSS } from 'lit';
import { styleMap } from 'lit/directives/style-map.js';
import type { LoadedPresentation, ElementRendererRegistry } from 'ooxml-ui/pptx/dom';
import css from './presentation-preview.css?raw';

/** An embedded reading surface over the shared PowerPoint DOM renderer. */
export class TeamsPresentationPreview extends LitElement {
	static override styles = unsafeCSS(css);
	static override properties = {
		data: { state: true },
		slide: { state: true },
		warnings: { state: true },
		availableWidth: { state: true },
		availableHeight: { state: true },
		fontCss: { state: true },
	};
	declare data: LoadedPresentation | null;
	declare slide: number;
	declare warnings: string[];
	declare availableWidth: number;
	declare availableHeight: number;
	declare fontCss: string;
	private renderer: typeof import('ooxml-ui/pptx/dom') | undefined;
	private registry: ElementRendererRegistry | undefined;
	private stageNode: HTMLElement | null = null;
	private observer: ResizeObserver | undefined;
	private resizeFrame = 0;
	private fontUrls: string[] = [];
	private generation = 0;

	constructor() {
		super();
		this.data = null;
		this.slide = 0;
		this.warnings = [];
		this.availableWidth = 0;
		this.availableHeight = 0;
		this.fontCss = '';
	}
	async load(bytes: Uint8Array): Promise<void> {
		const generation = ++this.generation;
		this.clear();
		const renderer = await import('ooxml-ui/pptx/dom');
		if (generation !== this.generation || !this.isConnected) return;
		const data = await renderer.loadPresentation(new Uint8Array(bytes).buffer, {
			allowExternalImages: false,
			maxUncompressedBytes: 134_217_728,
		});
		if (generation !== this.generation || !this.isConnected) {
			data.handler.dispose();
			renderer.revokeBlobUrls(data.blobUrls);
			return;
		}
		try {
			const { width, height } = data.canvasSize;
			if (!data.slides.length) throw new Error('The presentation has no slides');
			if (!(width > 0 && height > 0 && Number.isFinite(width) && Number.isFinite(height)))
				throw new Error('The presentation has an invalid slide size');
			this.renderer = renderer;
			this.registry = renderer.createDefaultRegistry();
			this.data = data;
			this.slide = 0;
			this.warnings = [
				...new Set(
					[...data.warnings, ...data.slides.flatMap((slide) => slide.warnings ?? [])].map(
						(warning) => warning.message,
					),
				),
			];
			const fonts = renderer.buildEmbeddedFontStyles(data.embeddedFonts, (bytes, mime) =>
				URL.createObjectURL(new Blob([new Uint8Array(bytes)], { type: mime })),
			);
			this.fontCss = fonts.fontFaceCss;
			this.fontUrls = fonts.objectUrls;
			renderer.glyphOutlineFontCache.registerEmbeddedFonts(data.embeddedFonts);
			this.showSlide();
			await this.updateComplete;
			if (generation !== this.generation || !this.isConnected) return;
			const stage = this.renderRoot.querySelector<HTMLElement>('.stage');
			if (stage && typeof ResizeObserver !== 'undefined') {
				this.observer = new ResizeObserver((entries) => {
					const bounds = entries[0]?.contentRect;
					if (!bounds || !bounds.width || !bounds.height) return;
					const { width, height } = bounds;
					if (
						Math.abs(width - this.availableWidth) < 0.5 &&
						Math.abs(height - this.availableHeight) < 0.5
					)
						return;
					cancelAnimationFrame(this.resizeFrame);
					this.resizeFrame = requestAnimationFrame(() => {
						if (generation === this.generation && this.isConnected) {
							this.availableWidth = width;
							this.availableHeight = height;
						}
					});
				});
				this.observer.observe(stage);
			}
		} catch (error) {
			if (this.data === data) this.clear();
			else {
				data.handler.dispose();
				renderer.revokeBlobUrls(data.blobUrls);
			}
			throw error;
		}
	}
	private stopMedia(): void {
		for (const media of this.stageNode?.querySelectorAll('video, audio') ?? []) {
			const player = media as HTMLMediaElement;
			player.pause();
			player.removeAttribute('src');
			player.load();
		}
		this.stageNode?.remove();
		this.stageNode = null;
	}
	private clear(): void {
		cancelAnimationFrame(this.resizeFrame);
		this.observer?.disconnect();
		this.observer = undefined;
		this.stopMedia();
		this.data?.handler.dispose();
		this.renderer?.revokeBlobUrls([...(this.data?.blobUrls ?? []), ...this.fontUrls]);
		this.fontUrls = [];
		this.fontCss = '';
		this.data = null;
		this.renderer = undefined;
		this.registry = undefined;
	}
	override disconnectedCallback(): void {
		this.generation++;
		this.clear();
		super.disconnectedCallback();
	}
	private showSlide(): void {
		const { data, renderer, registry } = this;
		const slide = data?.slides[this.slide];
		if (!data || !slide || !renderer || !registry) return;
		this.stopMedia();
		this.stageNode = renderer.renderSlideStage({
			document: this.ownerDocument,
			slide,
			canvasSize: data.canvasSize,
			mediaDataUrls: data.mediaDataUrls,
			registry,
			t: renderer.createTranslator(),
			...(data.colorScheme ? { colorScheme: data.colorScheme } : {}),
			...(data.fontScheme ? { fontScheme: data.fontScheme } : {}),
			...(data.tableStyleMap ? { tableStyleMap: data.tableStyleMap } : {}),
			fieldContext: renderer.buildFieldSubstitutionContext({
				headerFooter: data.headerFooter,
				customProperties: data.customProperties,
				slide,
			}),
			reading: true,
		});
		this.stageNode.setAttribute('aria-label', `Slide ${this.slide + 1}`);
	}
	private move(delta: number): void {
		const next = Math.max(0, Math.min((this.data?.slides.length ?? 1) - 1, this.slide + delta));
		if (next === this.slide) return;
		this.slide = next;
		this.showSlide();
	}
	private navigate(event: KeyboardEvent): void {
		if (event.composedPath()[0] !== event.currentTarget) return;
		if (['ArrowRight', 'PageDown', 'ArrowLeft', 'PageUp'].includes(event.key)) {
			event.preventDefault();
			this.move(['ArrowRight', 'PageDown'].includes(event.key) ? 1 : -1);
		}
	}
	protected override render() {
		const data = this.data;
		if (!data) return nothing;
		const slide = data.slides[this.slide]!;
		const scale = Math.min(
			(this.availableWidth || data.canvasSize.width) / data.canvasSize.width,
			(this.availableHeight || data.canvasSize.height) / data.canvasSize.height,
		);
		const width = data.canvasSize.width * scale;
		if (this.stageNode) this.stageNode.style.transform = `scale(${scale})`;
		return html`<style>
				${this.fontCss}
			</style>
			<p class="hint">
				Reading view. Animation, transitions and editing are unavailable. Layout and fonts may
				differ from PowerPoint.
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
				<div
					class="stage-frame"
					tabindex="0"
					role="group"
					aria-label="Slide navigation canvas"
					@keydown=${this.navigate}
					style=${styleMap({ width: `${width}px`, height: `${data.canvasSize.height * scale}px` })}
				>
					${this.stageNode}
				</div>
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
