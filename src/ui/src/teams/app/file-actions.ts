import { LitElement, html, nothing, unsafeCSS, type PropertyValues } from 'lit';
import { contentUrl, readContent, type FileEntry, type TeamsClient } from 'ooxml-core/teams';
import { downloadBytes } from '../../xlsx/file-commands.js';
import css from './file-actions.css?raw';
import { toggleAnchoredPopover } from './anchored-popover.js';
const fileKey = (file: FileEntry | null | undefined): string =>
	JSON.stringify([file?.channelId, file?.messageId, file?.url, file?.name]);

export class TeamsFileActions extends LitElement {
	static override styles = unsafeCSS(css);
	static override properties = {
		file: { attribute: false },
		client: { attribute: false },
		busy: { state: true },
		error: { state: true },
		copied: { state: true },
		copyUrl: { state: true },
		menuOpen: { state: true },
	};
	declare file: FileEntry | null;
	declare client: TeamsClient | null;
	declare busy: boolean;
	declare error: string;
	declare copied: boolean;
	declare copyUrl: string;
	declare menuOpen: boolean;
	private abort: AbortController | undefined;
	private toggleMenu(): void {
		const menu = this.renderRoot.querySelector<HTMLElement>('[popover]')!;
		toggleAnchoredPopover(menu, this.renderRoot.querySelector('button')!);
	}
	constructor() {
		super();
		this.file = null;
		this.client = null;
		this.busy = false;
		this.error = '';
		this.copied = false;
		this.copyUrl = '';
		this.menuOpen = false;
	}
	override disconnectedCallback(): void {
		this.abort?.abort();
		super.disconnectedCallback();
	}
	protected override willUpdate(changed: PropertyValues<this>): void {
		if (
			(changed.has('file') && fileKey(changed.get('file')) !== fileKey(this.file)) ||
			(changed.has('client') && changed.get('client') !== undefined)
		) {
			this.abort?.abort();
			this.copyUrl = '';
			this.error = '';
			this.copied = false;
			this.renderRoot.querySelector<HTMLElement>('[popover]')?.hidePopover();
		}
	}
	private open(action: 'open' | 'browser' | 'pin'): void {
		this.renderRoot.querySelector<HTMLElement>('[popover]')?.hidePopover();
		this.dispatchEvent(
			new CustomEvent(`teams-files-${action}`, {
				detail: { attachment: this.file, ...(action === 'open' ? { mode: 'teams' } : {}) },
				bubbles: true,
				composed: true,
			}),
		);
	}
	private async transfer(action: 'copy' | 'download'): Promise<void> {
		if (this.busy || !this.file || !this.client) return;
		const { file, client } = this;
		const abort = new AbortController();
		this.abort = abort;
		this.busy = true;
		this.error = '';
		this.copied = false;
		try {
			const url = contentUrl(await client.fileUrl(file), this.ownerDocument.baseURI);
			if (!this.isConnected || abort.signal.aborted) return;
			if (!url) throw new Error('The file is no longer available');
			if (action === 'download') {
				const bytes = await readContent(url, abort.signal, 33_554_432);
				if (this.isConnected && !abort.signal.aborted)
					downloadBytes(
						this.ownerDocument,
						bytes,
						file.name,
						file.mime ?? 'application/octet-stream',
					);
			} else {
				this.copyUrl = url;
				try {
					await navigator.clipboard.writeText(url);
					if (!abort.signal.aborted && this.isConnected) this.copied = true;
				} catch {
					/* The selectable URL remains available when clipboard access is blocked. */
				}
			}
		} catch (error) {
			if (!abort.signal.aborted)
				this.error = error instanceof Error ? error.message : 'Could not access this file';
		} finally {
			if (this.abort === abort) {
				this.abort = undefined;
				this.busy = false;
			}
		}
	}
	protected override render() {
		return html`<button
				type="button"
				aria-label=${`More actions for ${this.file?.name ?? 'file'}`}
				?disabled=${!this.file?.url}
				aria-controls="actions"
				aria-expanded=${String(this.menuOpen)}
				@click=${this.toggleMenu}
			>
				…
			</button>
			<div
				id="actions"
				popover="auto"
				aria-label="File actions"
				@toggle=${() => (this.menuOpen = this.renderRoot.querySelector<HTMLElement>('[popover]')!.matches(':popover-open'))}
			>
				<button type="button" @click=${() => this.open('open')}>Open in OpenTeams</button>
				<button type="button" @click=${() => this.open('browser')}>Open in browser</button>
				<button type="button" ?disabled=${this.busy} @click=${() => void this.transfer('download')}>
					Download
				</button>
				<button type="button" ?disabled=${this.busy} @click=${() => void this.transfer('copy')}>
					Copy link
				</button>
				<button type="button" @click=${() => this.open('pin')}>Pin as tab</button>
				${
					this.busy
						? html`<p role="status">Preparing file…</p>
								<button type="button" @click=${() => this.abort?.abort()}>Cancel</button>`
						: nothing
				}
				${this.error ? html`<p role="alert">${this.error}</p>` : nothing}
				${this.copied ? html`<p role="status">Link copied</p>` : nothing}
				${this.copyUrl ? html`<label>File link<input readonly .value=${this.copyUrl} @focus=${(event: Event) => (event.target as HTMLInputElement).select()} /></label>` : nothing}
			</div>`;
	}
}
export function defineTeamsFileActions(): void {
	if (!customElements.get('teams-file-actions'))
		customElements.define('teams-file-actions', TeamsFileActions);
}
