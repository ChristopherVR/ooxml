import { LitElement, html, nothing, unsafeCSS, type PropertyValues } from 'lit';
import { keyed } from 'lit/directives/keyed.js';
import {
	contentUrl,
	detectContentKind,
	markdownBlocks,
	markdownInline,
	readContent,
	type ContentKind,
	type OfficeKind,
} from 'ooxml-core/teams';
import type { OpenFileDetail } from './teams-app.js';
import css from './content-preview.css?raw';

/** Trusted host adapters return an embedding page URL, not an Office file URL. */
export type FileEmbeds = Partial<Record<OfficeKind, (detail: OpenFileDetail) => string>>;

export class TeamsContentPreview extends LitElement {
	static override styles = unsafeCSS(css);
	static override properties = {
		detail: { attribute: false },
		embeds: { attribute: false },
		status: { state: true },
		text: { state: true },
		frame: { state: true },
		native: { state: true },
		error: { state: true },
	};
	declare detail: OpenFileDetail | null;
	declare embeds: FileEmbeds;
	declare status: 'loading' | 'ready' | 'error';
	declare text: string;
	declare frame: string | null;
	declare native: ContentKind | null;
	declare error: string;
	private abort: AbortController | undefined;

	constructor() {
		super();
		this.detail = null;
		this.embeds = {};
		this.status = 'loading';
		this.text = '';
		this.frame = null;
		this.native = null;
		this.error = '';
	}

	protected override updated(changed: PropertyValues<this>): void {
		if (changed.has('detail') || changed.has('embeds')) void this.load();
	}

	override disconnectedCallback(): void {
		this.abort?.abort();
		super.disconnectedCallback();
	}

	override connectedCallback(): void {
		super.connectedCallback();
		if (this.hasUpdated) void this.load();
	}

	private async load(): Promise<void> {
		this.abort?.abort();
		const request = (this.abort = new AbortController());
		this.status = 'loading';
		this.text = '';
		this.frame = null;
		this.native = null;
		const detail = this.detail;
		if (!detail) return;
		const kind = detectContentKind(detail.attachment.name, detail.attachment.mime);
		try {
			const url = contentUrl(detail.url, this.ownerDocument.baseURI);
			if (!url) throw new Error('This content has no valid web URL');
			const embed = this.embeds[detail.attachment.kind];
			if (embed || kind === 'website') {
				this.frame = contentUrl(embed ? embed(detail) : url, this.ownerDocument.baseURI);
				if (!this.frame) throw new Error('The viewer returned an invalid embed URL');
				this.status = 'ready';
				return;
			}
			if (kind === 'pptx' || kind === 'other')
				throw new Error('No embedded viewer is configured for this file type');
			const bytes = await readContent(
				url,
				request.signal,
				kind === 'markdown' || kind === 'text' ? 2_097_152 : 33_554_432,
			);
			if (request.signal.aborted) return;
			if (kind === 'markdown' || kind === 'text') {
				this.text = new TextDecoder().decode(bytes);
			} else {
				if (kind === 'docx') (await import('../../docx/index.js')).registerDocxEditor();
				if (kind === 'xlsx') (await import('../../xlsx/index.js')).defineXlsxEditor();
				if (kind === 'vsdx') (await import('../../visio/index.js')).registerVisioViewer();
				if (request.signal.aborted) return;
				this.native = kind;
				await this.updateComplete;
				if (request.signal.aborted) return;
				const viewer = this.shadowRoot?.querySelector<
					HTMLElement & { load(bytes: Uint8Array, name?: string): Promise<void> }
				>('docx-editor, xlsx-editor, visio-viewer');
				if (!viewer) throw new Error('The viewer could not be mounted');
				await viewer.load(bytes, detail.attachment.name);
			}
			if (!request.signal.aborted) this.status = 'ready';
		} catch (error) {
			if (request.signal.aborted) return;
			this.native = null;
			this.error = error instanceof Error ? error.message : 'Could not preview content';
			this.status = 'error';
		}
	}

	protected override render() {
		const detail = this.detail;
		if (!detail) return nothing;
		const kind = detectContentKind(detail.attachment.name, detail.attachment.mime);
		const url = contentUrl(detail.url, this.ownerDocument.baseURI);
		return html`<section aria-label="Content preview">
			<header>
				<h2>${detail.attachment.name}</h2>
				${url ? html`<a href=${url} target="_blank" rel="noopener noreferrer">Open externally</a>` : nothing}
				<button
					type="button"
					@click=${() => this.dispatchEvent(new CustomEvent('teams-preview-close', { bubbles: true, composed: true }))}
				>
					Close preview
				</button>
			</header>
			${this.status === 'loading' ? html`<p role="status">Loading preview...</p>` : nothing}
			${
				this.status === 'error'
					? html`<p role="alert">${this.error}</p>
							<button type="button" @click=${() => void this.load()}>Retry</button>`
					: nothing
			}
			${
				this.frame
					? html`<p class="hint">If this site blocks embedding, use Open externally.</p>
							<iframe
								title=${detail.attachment.name}
								src=${this.frame}
								sandbox="allow-scripts allow-forms"
								referrerpolicy="no-referrer"
							></iframe>`
					: nothing
			}
			${keyed(
				detail,
				this.native === 'docx'
					? html`<docx-editor .readOnly=${true} .fileName=${detail.attachment.name}></docx-editor>`
					: this.native === 'xlsx'
						? html`<xlsx-editor
								.readOnly=${true}
								.fileName=${detail.attachment.name}
							></xlsx-editor>`
						: this.native === 'vsdx'
							? html`<visio-viewer .showToolbar=${false}></visio-viewer>`
							: nothing,
			)}
			${
				this.status === 'ready' && !this.frame && !this.native
					? html`<article>
							${
								kind === 'markdown'
									? markdownBlocks(this.text).map((block) => {
											if (block.kind === 'code') return html`<pre><code>${block.text}</code></pre>`;
											const text = this.inline(block.text, url!);
											if (block.kind === 'quote') return html`<blockquote>${text}</blockquote>`;
											if (block.kind === 'list')
												return html`<ul>
													<li>${text}</li>
												</ul>`;
											if (block.kind === 'heading')
												return html`<div role="heading" aria-level=${block.level} class="heading">
													${text}
												</div>`;
											return html`<p>${text}</p>`;
										})
									: html`<pre>${this.text}</pre>`
							}
						</article>`
					: nothing
			}
		</section>`;
	}

	private inline(text: string, base: string) {
		return markdownInline(text, base).map((token) => {
			if (token.kind === 'strong') return html`<strong>${token.text}</strong>`;
			if (token.kind === 'emphasis') return html`<em>${token.text}</em>`;
			if (token.kind === 'code') return html`<code>${token.text}</code>`;
			if (token.kind === 'link')
				return html`<a href=${token.url!} target="_blank" rel="noopener noreferrer"
					>${token.text}</a
				>`;
			return token.text;
		});
	}
}

export function defineTeamsContentPreview(): void {
	if (globalThis.customElements && !customElements.get('teams-content-preview'))
		customElements.define('teams-content-preview', TeamsContentPreview);
}
