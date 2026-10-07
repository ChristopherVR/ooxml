import { LitElement, html, nothing, unsafeCSS, type PropertyValues } from 'lit';
import { filterFiles, type Attachment, type FileEntry, type TeamsClient } from 'ooxml-core/teams';
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
	private retry: (() => Promise<unknown>) | undefined;

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
	}
	protected override updated(changed: PropertyValues<this>): void {
		if (changed.has('client')) {
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
	private async run(task: () => Promise<unknown>): Promise<void> {
		if (this.busy) return;
		const client = this.client;
		if (this.retry !== task) this.destination = this.channelName;
		this.busy = true;
		this.error = '';
		this.retry = task;
		try {
			await task();
			if (this.client === client) this.retry = undefined;
		} catch (error) {
			if (this.client === client)
				this.error = error instanceof Error ? error.message : 'Could not share files';
		} finally {
			this.busy = false;
		}
	}
	private upload(event: Event): void {
		const input = event.target as HTMLInputElement;
		const files = Array.from(input.files ?? []);
		input.value = '';
		const { client, channelId } = this;
		if (!client || !files.length) return;
		void this.run(() => client.uploadFiles(channelId, files));
	}
	private create(event: SubmitEvent): void {
		event.preventDefault();
		const name = String(new FormData(event.target as HTMLFormElement).get('name') ?? '');
		const { client, channelId } = this;
		if (!client) return;
		void this.run(async () => {
			const attachment = await client.createWorkbook(channelId, name);
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
			${this.busy ? html`<p role="status">Sharing files in # ${this.destination}...</p>` : nothing}
			${
				this.error
					? html`<p role="alert">Sharing in # ${this.destination} failed: ${this.error}</p>
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
