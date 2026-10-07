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
	type FileOperationOptions,
} from 'ooxml-core/teams';
import type { OpenFileDetail } from './teams-app.js';
import css from './content-preview.css?raw';
import { blobFor, downloadBytes, saveExtension, withExtension } from '../../xlsx/file-commands.js';
import { workbookActions } from './workbook-actions.js';

/** Trusted host adapters return an embedding page URL, not an Office file URL. */
export type FileEmbeds = Partial<Record<OfficeKind, (detail: OpenFileDetail) => string>>;
export type SaveFileCopy = (file: File, options?: FileOperationOptions) => Promise<void>;

export class TeamsContentPreview extends LitElement {
	static override styles = unsafeCSS(css);
	static override properties = {
		detail: { attribute: false },
		embeds: { attribute: false },
		saveCopy: { attribute: false },
		editing: { state: true },
		dirty: { state: true },
		saving: { state: true },
		saveError: { state: true },
		saved: { state: true },
		downloaded: { state: true },
		status: { state: true },
		text: { state: true },
		frame: { state: true },
		native: { state: true },
		error: { state: true },
	};
	declare detail: OpenFileDetail | null;
	declare embeds: FileEmbeds;
	declare saveCopy: SaveFileCopy | undefined;
	declare editing: boolean;
	declare dirty: boolean;
	declare saving: boolean;
	declare saveError: string;
	declare saved: boolean;
	declare downloaded: boolean;
	private revision = 0;
	private saveAbort: AbortController | undefined;
	private readonly beforeUnload = (event: BeforeUnloadEvent): void => {
		this.commitPendingEdit();
		if (this.dirty || this.saving) {
			event.preventDefault();
			event.returnValue = '';
		}
	};
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
		this.saveCopy = undefined;
		this.editing = false;
		this.dirty = false;
		this.saving = false;
		this.saveError = '';
		this.saved = false;
		this.downloaded = false;
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
		this.saveAbort?.abort();
		this.abort?.abort();
		this.ownerDocument.defaultView?.removeEventListener('beforeunload', this.beforeUnload);
		super.disconnectedCallback();
	}

	override connectedCallback(): void {
		super.connectedCallback();
		this.ownerDocument.defaultView?.addEventListener('beforeunload', this.beforeUnload);
		if (this.hasUpdated) void this.load();
	}

	/** Called by the workspace before replacing this pane. */
	canLeave(): boolean {
		if (this.saving) {
			this.saveError = 'Wait for the workbook copy to finish saving or cancel the save.';
			return false;
		}
		if (!this.commitPendingEdit()) {
			this.saveError = 'Finish or cancel the current cell edit before leaving.';
			return false;
		}
		return !this.dirty || globalThis.confirm?.('Discard unsaved workbook edits?') === true;
	}

	private commitPendingEdit(): boolean {
		const viewer = this.shadowRoot?.querySelector<HTMLElement & { commitEdit(): boolean }>(
			'xlsx-editor',
		);
		return viewer?.commitEdit() ?? true;
	}

	private async load(): Promise<void> {
		this.saveAbort?.abort();
		this.saveAbort = undefined;
		this.abort?.abort();
		const request = (this.abort = new AbortController());
		this.editing = false;
		this.dirty = false;
		this.saving = false;
		this.saveError = '';
		this.saved = false;
		this.downloaded = false;
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
			if (kind === 'other') throw new Error('No embedded viewer is configured for this file type');
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
				if (kind === 'pptx')
					(await import('./presentation-preview.js')).defineTeamsPresentationPreview();
				if (request.signal.aborted) return;
				this.native = kind;
				await this.updateComplete;
				if (request.signal.aborted) return;
				const viewer = this.shadowRoot?.querySelector<
					HTMLElement & { load(bytes: Uint8Array, name?: string): Promise<void> }
				>('docx-editor, xlsx-editor, visio-viewer, teams-presentation-preview');
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

	private toggleWorkbook(): void {
		if (!this.commitPendingEdit()) {
			this.saveError = 'Finish or cancel the current cell edit before switching modes.';
			return;
		}
		this.saveError = '';
		this.editing = !this.editing;
	}
	private async saveWorkbookCopy(download = false): Promise<void> {
		if (this.saving || (!download && !this.saveCopy) || this.native !== 'xlsx' || !this.detail)
			return;
		if (!this.commitPendingEdit()) {
			this.saveError = 'Finish or cancel the current cell edit before saving.';
			return;
		}
		const viewer = this.shadowRoot?.querySelector<
			HTMLElement & { saveBytes(): Promise<Uint8Array>; markClean(): void }
		>('xlsx-editor');
		if (!viewer) return;
		const callback = this.saveCopy;
		const request = this.abort;
		const controller = new AbortController();
		this.saveAbort = controller;
		const name = withExtension(
			this.detail.attachment.name,
			saveExtension(this.detail.attachment.name),
		);
		this.saving = true;
		this.saveError = '';
		this.saved = false;
		this.downloaded = false;
		try {
			const pendingBytes = viewer.saveBytes();
			const revision = this.revision;
			const bytes = await pendingBytes;
			if (request?.signal.aborted || controller.signal.aborted) return;
			if (download) {
				downloadBytes(this.ownerDocument, bytes, name);
				this.downloaded = true;
				return;
			}
			const blob = blobFor(bytes, name);
			await callback!(new File([blob], name, { type: blob.type }), { signal: controller.signal });
			if (request?.signal.aborted || controller.signal.aborted) return;
			if (revision === this.revision) viewer.markClean();
			this.saveError = '';
			this.saved = true;
		} catch (error) {
			if (!request?.signal.aborted && !controller.signal.aborted)
				this.saveError =
					error instanceof Error ? error.message : 'Could not save the workbook copy';
		} finally {
			if (this.saveAbort === controller) {
				this.saveAbort = undefined;
				if (!request?.signal.aborted) this.saving = false;
			}
		}
	}

	private cancelWorkbookSave(): void {
		this.saveAbort?.abort();
		this.saveAbort = undefined;
		this.saving = false;
		this.saveError = 'Workbook save canceled. Your edits remain local.';
	}

	protected override render() {
		const detail = this.detail;
		if (!detail) return nothing;
		const kind = detectContentKind(detail.attachment.name, detail.attachment.mime);
		const url = contentUrl(detail.url, this.ownerDocument.baseURI);
		return html`<section aria-label="Content preview">
			<header>
				<h2>${detail.attachment.name}</h2>
				${
					this.native === 'xlsx' && this.status === 'ready'
						? workbookActions(
								{ editing: this.editing, saving: this.saving, canShare: !!this.saveCopy },
								{
									toggle: () => this.toggleWorkbook(),
									share: () => void this.saveWorkbookCopy(),
									download: () => void this.saveWorkbookCopy(true),
									cancel: () => this.cancelWorkbookSave(),
								},
							)
						: nothing
				}
				${url ? html`<a href=${url} target="_blank" rel="noopener noreferrer">Open externally</a>` : nothing}
				<button
					type="button"
					@click=${() => this.dispatchEvent(new CustomEvent('teams-preview-close', { bubbles: true, composed: true }))}
				>
					Close preview
				</button>
			</header>
			${this.native === 'xlsx' ? html`<p class="hint">Edits stay local. Download a copy${this.saveCopy ? ' or save a copy to the channel' : ''} to keep them. Other people keep their own copy open.</p>` : nothing}
			${this.saved ? html`<p role="status">Workbook copy shared in the channel.</p>` : nothing}
			${this.downloaded ? html`<p role="status">Workbook download started. Your changes remain local.</p>` : nothing}
			${this.saveError ? html`<p role="alert">${this.saveError}</p>` : nothing}
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
								.readOnly=${!this.editing}
								.fileName=${detail.attachment.name}
								@workbook-change=${() => {
									this.revision++;
									this.saved = false;
									this.downloaded = false;
								}}
								@dirty-change=${(e: CustomEvent<{ dirty: boolean }>) => (this.dirty = e.detail.dirty)}
								@readonly-change=${(e: CustomEvent<{ readOnly: boolean }>) => {
									this.editing = !e.detail.readOnly;
								}}
								@file-command=${(e: CustomEvent<{ command: string }>) => {
									if (
										['save', 'saveAs', 'export', 'exportCsv'].includes(e.detail.command) &&
										!this.commitPendingEdit()
									) {
										e.preventDefault();
										this.saveError = 'Finish or cancel the current cell edit before saving.';
										return;
									}
									if (e.detail.command === 'save') {
										e.preventDefault();
										void this.saveWorkbookCopy(!this.saveCopy);
									}
									if (e.detail.command === 'new' || e.detail.command === 'open') e.preventDefault();
								}}
							></xlsx-editor>`
						: this.native === 'vsdx'
							? html`<visio-viewer .showToolbar=${false}></visio-viewer>`
							: this.native === 'pptx'
								? html`<teams-presentation-preview></teams-presentation-preview>`
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
