import { LitElement, html, nothing, unsafeCSS, type PropertyValues } from 'lit';
import {
	filterFiles,
	type Attachment,
	type FileEntry,
	type TeamsClient,
	type FileOperationOptions,
	type FileTransferProgress,
} from 'ooxml-core/teams';
import css from './files-panel.css?raw';

/** File controls bind to channel-capturing core actions; uploaded bytes never enter UI state. */
export class TeamsFilesPanel extends LitElement {
	static override styles = unsafeCSS(css);
	static override properties = {
		client: { attribute: false },
		files: { attribute: false },
		channelId: { attribute: false },
		channelName: { attribute: false },
		canUpload: { type: Boolean },
		query: { state: true },
		creating: { state: true },
		busy: { state: true },
		error: { state: true },
		destination: { state: true },
		transfer: { state: true },
		canceled: { state: true },
	};
	declare client: TeamsClient | null;
	declare files: FileEntry[];
	declare channelId: string;
	declare channelName: string;
	declare canUpload: boolean;
	declare query: string;
	declare creating: boolean;
	declare busy: boolean;
	declare error: string;
	declare destination: string;
	declare transfer: FileTransferProgress | null;
	declare canceled: boolean;
	private active: AbortController | undefined;
	private retry: ((options: FileOperationOptions) => Promise<unknown>) | undefined;

	constructor() {
		super();
		this.client = null;
		this.files = [];
		this.channelId = '';
		this.channelName = '';
		this.canUpload = false;
		this.query = '';
		this.creating = false;
		this.busy = false;
		this.error = '';
		this.destination = '';
		this.transfer = null;
		this.canceled = false;
	}
	override disconnectedCallback(): void {
		this.active?.abort();
		super.disconnectedCallback();
	}
	protected override updated(changed: PropertyValues<this>): void {
		if (changed.has('client')) {
			this.active?.abort();
			this.active = undefined;
			this.busy = false;
			this.transfer = null;
			this.canceled = false;
			this.retry = undefined;
			this.error = '';
		}
	}
	private emit(
		action: 'open' | 'pin',
		attachment: FileEntry | (Attachment & { channelId: string }),
	): void {
		this.dispatchEvent(
			new CustomEvent(`teams-files-${action}`, {
				detail: { attachment },
				bubbles: true,
				composed: true,
			}),
		);
	}
	private async run(task: (options: FileOperationOptions) => Promise<unknown>): Promise<void> {
		if (this.busy) return;
		const client = this.client;
		const controller = new AbortController();
		this.active = controller;
		if (this.retry !== task) this.destination = this.channelName;
		this.busy = true;
		this.error = '';
		this.canceled = false;
		this.transfer = null;
		this.retry = task;
		try {
			await task({
				signal: controller.signal,
				onProgress: (progress) => {
					if (this.active === controller && this.client === client && this.isConnected)
						this.transfer = progress;
				},
			});
			if (this.client === client && this.active === controller) this.retry = undefined;
		} catch (error) {
			if (this.client === client && this.active === controller && this.isConnected) {
				this.canceled = controller.signal.aborted;
				if (!this.canceled)
					this.error = error instanceof Error ? error.message : 'Could not share files';
			}
		} finally {
			if (this.active === controller) {
				this.active = undefined;
				this.busy = false;
			}
		}
	}
	private upload(event: Event): void {
		const input = event.target as HTMLInputElement;
		const files = Array.from(input.files ?? []);
		input.value = '';
		const { client, channelId } = this;
		if (!client || !files.length) return;
		void this.run((options) => client.uploadFiles(channelId, files, options));
	}
	private create(event: SubmitEvent): void {
		event.preventDefault();
		const name = String(new FormData(event.target as HTMLFormElement).get('name') ?? '');
		const { client, channelId } = this;
		if (!client) return;
		void this.run(async (options) => {
			const attachment = await client.createWorkbook(channelId, name, options);
			if (!this.isConnected || this.client !== client || this.channelId !== channelId) return;
			this.creating = false;
			this.emit('open', { ...attachment, channelId });
		});
	}
	protected override render() {
		const files = filterFiles(this.files, this.query);
		const disabled = this.busy || !this.canUpload || !this.channelId;
		return html`<div class="toolbar">
				<input
					type="search"
					aria-label="Search files"
					placeholder="Search files"
					.value=${this.query}
					@input=${(e: Event) => (this.query = (e.target as HTMLInputElement).value)}
				/>
				<label
					>Upload files
					<input
						type="file"
						aria-label="Upload files"
						multiple
						?disabled=${disabled}
						@change=${this.upload}
				/></label>
				<button
					type="button"
					?disabled=${disabled}
					@click=${() => (this.creating = !this.creating)}
				>
					New Excel workbook
				</button>
			</div>
			<p class="hint">
				${this.canUpload ? `New files are shared in # ${this.channelName}.` : 'Configure file storage to upload files or create workbooks.'}
			</p>
			${
				this.creating
					? html`<form aria-label="Create Excel workbook" @submit=${this.create}>
							<input
								name="name"
								aria-label="Workbook name"
								placeholder="Workbook name"
								maxlength="75"
								required
								?disabled=${this.busy}
							/>
							<button type="submit" ?disabled=${disabled}>Create workbook</button>
							<button type="button" ?disabled=${this.busy} @click=${() => (this.creating = false)}>
								Cancel
							</button>
						</form>`
					: nothing
			}
			${
				this.busy
					? html`<div class="transfer">
							<p role="status">
								Sharing files in # ${this.destination}:
								${this.transfer?.phase === 'preparing' ? 'Preparing' : 'Uploading'}
								${this.transfer?.fileName ?? ''} (${this.transfer?.completed ?? 0} of
								${this.transfer?.total ?? 0} uploaded)
							</p>
							${this.transfer ? html`<progress aria-label="Files uploaded" max=${this.transfer.total} value=${this.transfer.completed}></progress>` : nothing}
							<button type="button" @click=${() => this.active?.abort()}>Cancel sharing</button>
						</div>`
					: nothing
			}
			${
				this.error || this.canceled
					? html`${this.canceled ? html`<p role="status">Sharing in # ${this.destination} canceled.</p>` : html`<p role="alert">Sharing in # ${this.destination} failed: ${this.error}</p>`}
							<button
								type="button"
								?disabled=${this.busy}
								@click=${() => this.retry && void this.run(this.retry)}
							>
								Retry sharing files
							</button>`
					: nothing
			}
			<ul class="files">
				${
					files.length
						? files.map(
								(file) => html`<li>
									<span class="badge" data-kind=${file.kind}
										>${file.kind === 'other' ? 'F' : file.kind[0]!.toUpperCase()}</span
									>
									<span class="meta"
										><strong>${file.name}</strong
										><small
											>${file.author} ·
											${new Date(file.ts).toLocaleString()}${file.channelName ? ` · # ${file.channelName}` : ''}</small
										></span
									>
									<button type="button" @click=${() => this.emit('open', file)}>Open</button>
									<button
										type="button"
										?disabled=${!file.url}
										@click=${() => this.emit('pin', file)}
									>
										Pin as tab
									</button>
								</li>`,
							)
						: html`<li class="none">
								${this.query.trim() ? 'No matching files.' : 'No files shared yet.'}
							</li>`
				}
			</ul>`;
	}
}

export function defineTeamsFilesPanel(): void {
	if (globalThis.customElements && !customElements.get('teams-files-panel'))
		customElements.define('teams-files-panel', TeamsFilesPanel);
}
