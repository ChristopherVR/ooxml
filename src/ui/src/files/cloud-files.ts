import { html, nothing } from 'lit';
import { ifDefined } from 'lit/directives/if-defined.js';
import type {
	CloudAccount,
	CloudEntry,
	CloudFileKind,
	CloudFileProvider,
	CloudListing,
	CloudOpenedFile,
} from 'ooxml-core/embedding';
import { OfficeElement, controlStyles } from '../base';
import { defineSearchField } from '../form/search-field';
import { definer } from '../registry';
import css from './cloud-files.css?raw';
import {
	DEFAULT_CLOUD_FILES_LABELS,
	entryBadge,
	fillLabel,
	formatModified,
	formatSize,
	messageOf,
	type OfficeCloudFilesLabels,
} from './cloud-files-labels';

export { DEFAULT_CLOUD_FILES_LABELS, type OfficeCloudFilesLabels } from './cloud-files-labels';

/** Detail of the bubbling, composed `office-cloud-file-open` event. */
export interface OfficeCloudFileOpenDetail {
	providerId: string;
	file: CloudOpenedFile;
}
export type OfficeCloudFileOpenEvent = CustomEvent<OfficeCloudFileOpenDetail>;

type View = 'recent' | 'files' | 'search';
interface Crumb {
	id: string | undefined;
	name: string;
}

const ALL_KINDS: readonly CloudFileKind[] = ['docx', 'xlsx', 'pptx', 'vsdx'];
const LINK_ID = '\u0000link';

/**
 * `<office-ui-cloud-files>`: browse and open files from one or more `CloudFileProvider`s (OneDrive,
 * or a host's own). Properties: `providers`, `accept` (default all four Office kinds; a file of
 * another kind is listed but `aria-disabled`, folders are always enterable), `labels` (partial,
 * English defaults), `searchDelay` (ms, default 300) and `locale` (dates). Opening a file emits the
 * bubbling, composed `office-cloud-file-open` event with `{ providerId, file }`; a failure shows
 * inline and emits nothing. In-flight listings are aborted when the user navigates away.
 */
export class OfficeUiCloudFiles extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		providers: { attribute: false },
		accept: { attribute: false },
		labels: { attribute: false },
		searchDelay: { type: Number, attribute: 'search-delay' },
		locale: { type: String },
		active: { state: true },
		view: { state: true },
		path: { state: true },
		entries: { state: true },
		next: { state: true },
		loading: { state: true },
		error: { state: true },
		query: { state: true },
		opening: { state: true },
		progress: { state: true },
		focusIndex: { state: true },
	};
	declare providers: readonly CloudFileProvider[];
	declare accept: readonly CloudFileKind[];
	declare labels: Partial<OfficeCloudFilesLabels>;
	declare searchDelay: number;
	declare locale: string | undefined;
	declare active: string | undefined;
	declare view: View;
	declare path: readonly Crumb[];
	declare entries: readonly CloudEntry[];
	declare next: string | undefined;
	declare loading: boolean;
	declare error: string;
	declare query: string;
	declare opening: string | undefined;
	declare progress: { loaded: number; total?: number } | undefined;
	declare focusIndex: number;

	private listAbort: AbortController | undefined;
	private openAbort: AbortController | undefined;
	private searchTimer: ReturnType<typeof setTimeout> | undefined;
	private accountKey = '';
	private subscriptions: (() => void)[] = [];

	constructor() {
		super();
		this.providers = [];
		this.accept = ALL_KINDS;
		this.labels = {};
		this.searchDelay = 300;
		this.locale = undefined;
		this.active = undefined;
		this.view = 'files';
		this.path = [];
		this.entries = [];
		this.next = undefined;
		this.loading = false;
		this.error = '';
		this.query = '';
		this.opening = undefined;
		this.progress = undefined;
		this.focusIndex = 0;
	}

	private get text(): OfficeCloudFilesLabels {
		return { ...DEFAULT_CLOUD_FILES_LABELS, ...this.labels };
	}

	/** The provider being browsed: the chosen one, else the first. */
	get provider(): CloudFileProvider | undefined {
		return this.providers.find((p) => p.id === this.active) ?? this.providers[0];
	}

	override connectedCallback(): void {
		super.connectedCallback();
		this.subscribe();
	}

	override disconnectedCallback(): void {
		super.disconnectedCallback();
		this.unsubscribe();
		this.listAbort?.abort();
		this.openAbort?.abort();
		clearTimeout(this.searchTimer);
		this.accountKey = '';
	}

	protected override willUpdate(changed: Map<PropertyKey, unknown>): void {
		if (changed.has('providers') && this.isConnected) this.subscribe();
		if (changed.has('providers') || changed.has('active')) this.syncAccount();
	}

	private subscribe(): void {
		this.unsubscribe();
		for (const p of this.providers)
			this.subscriptions.push(p.subscribe?.(() => this.onAccountChange(p)) ?? (() => {}));
	}

	private unsubscribe(): void {
		for (const off of this.subscriptions) off();
		this.subscriptions = [];
	}

	private onAccountChange(p: CloudFileProvider): void {
		if (p === this.provider) this.syncAccount();
		this.requestUpdate();
	}

	/** Reload from scratch when the browsed provider or its account changed. */
	private syncAccount(): void {
		const p = this.provider;
		const key = p ? `${p.id}\u0000${p.account()?.name ?? ''}` : '';
		if (key === this.accountKey) return;
		this.accountKey = key;
		this.listAbort?.abort();
		this.openAbort?.abort();
		this.entries = [];
		this.next = undefined;
		this.error = '';
		this.query = '';
		this.opening = undefined;
		this.loading = false;
		this.path = [{ id: undefined, name: this.text.myFiles }];
		this.view = p?.recent ? 'recent' : 'files';
		if (p?.account()) void this.load();
	}

	private async fetch(p: CloudFileProvider, signal: AbortSignal, page?: string) {
		const options = { signal, ...(page === undefined ? {} : { page }) };
		if (this.view === 'recent') return p.recent!(options);
		if (this.view === 'search') return p.search!(this.query.trim(), options);
		return p.list(this.path[this.path.length - 1]?.id, options);
	}

	private async load(append = false): Promise<void> {
		const p = this.provider;
		if (!p) return;
		this.listAbort?.abort();
		const controller = new AbortController();
		this.listAbort = controller;
		this.loading = true;
		this.error = '';
		if (!append) {
			this.entries = [];
			this.next = undefined;
			this.focusIndex = 0;
		}
		try {
			const listing: CloudListing = await this.fetch(
				p,
				controller.signal,
				append ? this.next : undefined,
			);
			if (controller.signal.aborted) return;
			this.entries = append ? [...this.entries, ...listing.entries] : listing.entries;
			this.next = listing.next;
		} catch (error) {
			if (!controller.signal.aborted) this.error = messageOf(error, this.text.failed);
		} finally {
			if (this.listAbort === controller) this.loading = false;
		}
	}

	private show(view: View): void {
		this.view = view;
		void this.load();
	}

	private enter(index: number): void {
		const root = this.path.slice(0, index + 1);
		this.path = root;
		this.show('files');
	}

	private async signIn(): Promise<void> {
		this.error = '';
		try {
			await this.provider?.signIn();
			this.syncAccount();
		} catch (error) {
			this.error = messageOf(error, this.text.failed);
		}
		this.requestUpdate();
	}

	private async signOut(): Promise<void> {
		try {
			await this.provider?.signOut?.();
			this.syncAccount();
		} catch (error) {
			this.error = messageOf(error, this.text.failed);
		}
		this.requestUpdate();
	}

	private onSearchInput(event: Event): void {
		this.query = (event.target as HTMLInputElement).value;
		clearTimeout(this.searchTimer);
		this.searchTimer = setTimeout(() => {
			if (this.query.trim()) this.show('search');
			else if (this.view === 'search') this.show(this.provider?.recent ? 'recent' : 'files');
		}, this.searchDelay);
	}

	private accepted(entry: CloudEntry): boolean {
		return (
			entry.type === 'folder' || (entry.kind !== undefined && this.accept.includes(entry.kind))
		);
	}

	/** Folders navigate; accepted files download and then emit `office-cloud-file-open`. */
	private activate(entry: CloudEntry): void {
		if (entry.type === 'folder') {
			this.path = this.view === 'files' ? [...this.path, entry] : [this.path[0]!, entry];
			return this.show('files');
		}
		if (this.opening || !this.accepted(entry)) return;
		void this.run(entry.id, (options) => this.provider!.open(entry, options));
	}

	private async openLink(link: string): Promise<void> {
		if (this.opening || !link.trim()) return;
		await this.run(LINK_ID, (options) => this.provider!.openLink!(link.trim(), options));
	}

	private async run(
		id: string,
		start: (options: {
			signal: AbortSignal;
			onProgress: (loaded: number, total?: number) => void;
		}) => Promise<CloudOpenedFile>,
	): Promise<void> {
		const controller = new AbortController();
		this.openAbort = controller;
		this.opening = id;
		this.progress = undefined;
		this.error = '';
		try {
			const file = await start({
				signal: controller.signal,
				onProgress: (loaded, total) => {
					if (!controller.signal.aborted)
						this.progress = total === undefined ? { loaded } : { loaded, total };
				},
			});
			if (controller.signal.aborted) return;
			this.fire<OfficeCloudFileOpenDetail>('office-cloud-file-open', {
				providerId: this.provider!.id,
				file,
			});
		} catch (error) {
			if (!controller.signal.aborted) this.error = messageOf(error, this.text.failed);
		} finally {
			if (this.openAbort === controller) {
				this.opening = undefined;
				this.progress = undefined;
			}
		}
	}

	private onKeydown(event: KeyboardEvent): void {
		const last = this.entries.length - 1;
		const move: Record<string, number> = {
			ArrowDown: this.focusIndex + 1,
			ArrowUp: this.focusIndex - 1,
			Home: 0,
			End: last,
		};
		if (event.key in move) {
			event.preventDefault();
			this.focusIndex = Math.max(0, Math.min(last, move[event.key]!));
			this.renderRoot.querySelector<HTMLElement>(`[data-index="${this.focusIndex}"]`)?.focus();
		} else if (event.key === 'Enter' || event.key === ' ') {
			const entry = this.entries[this.focusIndex];
			if (!entry) return;
			event.preventDefault();
			this.activate(entry);
		}
	}

	private row(entry: CloudEntry, index: number) {
		const { text } = this;
		const ok = this.accepted(entry);
		const busy = this.opening === entry.id;
		const blocked = entry.type === 'file' && (!ok || (this.opening !== undefined && !busy));
		const badge = entryBadge(entry);
		const place = this.view === 'files' ? undefined : entry.location;
		return html`<li
			class="row"
			role="option"
			data-index=${index}
			data-entry=${entry.id}
			tabindex=${index === this.focusIndex ? 0 : -1}
			aria-selected=${index === this.focusIndex ? 'true' : 'false'}
			aria-disabled=${blocked ? 'true' : 'false'}
			aria-busy=${busy ? 'true' : 'false'}
			title=${ifDefined(entry.type === 'file' && !ok ? text.unsupported : undefined)}
			@click=${() => {
				this.focusIndex = index;
				this.activate(entry);
			}}
			><span class="name"
				><span
					class="icon ${entry.type === 'folder' ? 'folder' : entry.kind ? entry.kind : 'other'}"
					aria-hidden="true"
					>${badge}</span
				><span class="title"
					><b>${entry.name}</b>${place ? html`<small>${place}</small>` : nothing}</span
				></span
			><span class="cell">${formatModified(entry.modified, this.locale)}</span
			><span class="cell size">${entry.type === 'folder' ? '' : formatSize(entry.size)}</span></li
		>`;
	}

	private list() {
		const { text } = this;
		if (this.loading && this.entries.length === 0)
			return html`<div aria-hidden="true"
				><div class="skeleton"></div><div class="skeleton"></div><div class="skeleton"></div
			></div>`;
		if (this.entries.length === 0 && !this.error)
			return html`<div class="empty"
				>${this.view === 'search' ? text.emptySearch : text.empty}</div
			>`;
		return html`
			<div class="head" aria-hidden="true"
				><span>${text.name}</span><span>${text.modified}</span
				><span class="size">${text.size}</span></div
			>
			<ul role="listbox" aria-label=${text.files} @keydown=${this.onKeydown}
				>${this.entries.map((entry, index) => this.row(entry, index))}</ul
			>
			${
				this.next
					? html`<button
							class="more"
							type="button"
							?disabled=${this.loading}
							@click=${() => void this.load(true)}
							>${text.loadMore}</button
						>`
					: nothing
			}
		`;
	}

	private signedOut(p: CloudFileProvider) {
		const { text } = this;
		const values = { provider: p.label };
		return html`<div class="signin"
			><h2>${fillLabel(text.signInHeading, values)}</h2><p>${fillLabel(text.signInText, values)}</p
			><button type="button" class="primary" @click=${() => void this.signIn()}
				>${text.signIn}</button
			><div class="error" role="alert">${this.error}</div></div
		>`;
	}

	private toolbar(p: CloudFileProvider, account: CloudAccount) {
		const { text } = this;
		const tab = (view: View, label: string) =>
			html`<button
				type="button"
				data-view=${view}
				aria-pressed=${this.view === view ? 'true' : 'false'}
				@click=${() => this.show(view)}
				>${label}</button
			>`;
		return html`<div class="bar"
			>${
				this.providers.length > 1
					? html`<div class="tabs" role="group" aria-label=${text.providers}
							>${this.providers.map(
								(other) =>
									html`<button
										type="button"
										data-provider=${other.id}
										aria-pressed=${other === p ? 'true' : 'false'}
										@click=${() => (this.active = other.id)}
										>${other.label}</button
									>`,
							)}</div
						>`
					: nothing
			}
			<div class="tabs"
				>${p.recent ? tab('recent', text.recent) : nothing}${tab('files', text.myFiles)}</div
			>
			${
				p.search
					? html`<office-ui-search
							placeholder=${text.searchPlaceholder}
							.value=${this.query}
							@input=${this.onSearchInput}
						></office-ui-search>`
					: nothing
			}
			<span class="grow"></span><span class="account">${account.name}</span> ${
				p.signOut
					? html`<button type="button" data-action="sign-out" @click=${() => void this.signOut()}
							>${text.signOut}</button
						>`
					: nothing
			}</div
		>`;
	}

	private linkField() {
		const { text } = this;
		const submit = (event: Event) => {
			event.preventDefault();
			const input = this.renderRoot.querySelector<HTMLInputElement>('.link input');
			if (input) void this.openLink(input.value);
		};
		return html`<form class="bar link" @submit=${submit}
			><input
				type="text"
				aria-label=${text.openLink}
				placeholder=${text.openLinkPlaceholder}
			/><button type="submit" ?disabled=${this.opening !== undefined}
				>${text.openLinkButton}</button
			></form
		>`;
	}

	private statusLine() {
		const { text } = this;
		if (!this.opening) return nothing;
		const entry = this.entries.find((candidate) => candidate.id === this.opening);
		const name = entry?.name ?? text.openLink;
		const { progress } = this;
		const percent =
			progress?.total && progress.total > 0
				? Math.min(100, Math.round((progress.loaded / progress.total) * 100))
				: undefined;
		const message =
			percent === undefined
				? fillLabel(text.opening, { name })
				: fillLabel(text.openingPercent, { name, percent: String(percent) });
		return html`<div class="status"
			><span role="status" data-progress=${ifDefined(percent)}>${message}</span
			><button type="button" data-action="cancel" @click=${() => this.openAbort?.abort()}
				>${text.cancel}</button
			></div
		>`;
	}

	private crumbs() {
		return html`<nav aria-label=${this.text.location}
			><ol
				>${this.path.map(
					(crumb, index) =>
						html`<li
							><button
								type="button"
								data-crumb=${index}
								aria-current=${ifDefined(
									this.view === 'files' && index === this.path.length - 1 ? 'page' : undefined,
								)}
								@click=${() => this.enter(index)}
								>${index === 0 ? this.text.myFiles : crumb.name}</button
							></li
						>`,
				)}</ol
			></nav
		>`;
	}

	protected override render() {
		const p = this.provider;
		if (!p) return nothing;
		const account = p.account();
		if (!account) return this.signedOut(p);
		return html`<div aria-busy=${this.loading ? 'true' : 'false'}
			>${this.toolbar(p, account)}${p.openLink ? this.linkField() : nothing}${
				this.view === 'files' ? this.crumbs() : nothing
			}${this.list()}${this.statusLine()} <div class="error" role="alert">${this.error}</div></div
		>`;
	}
}

export const defineCloudFiles = definer('office-ui-cloud-files', () => OfficeUiCloudFiles, [
	defineSearchField,
]);
