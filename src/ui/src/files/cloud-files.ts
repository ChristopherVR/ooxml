import { nothing } from 'lit';
import type {
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
	messageOf,
	type OfficeCloudFilesLabels,
} from './cloud-files-labels';
import {
	renderBrowser,
	renderSignedOut,
	type CloudFilesCrumb,
	type CloudFilesView,
	type CloudFilesViewModel,
} from './cloud-files-view';

export { DEFAULT_CLOUD_FILES_LABELS, type OfficeCloudFilesLabels } from './cloud-files-labels';

/** Detail of the bubbling, composed `office-cloud-file-open` event. */
export interface OfficeCloudFileOpenDetail {
	providerId: string;
	file: CloudOpenedFile;
}
export type OfficeCloudFileOpenEvent = CustomEvent<OfficeCloudFileOpenDetail>;

type View = CloudFilesView;
type Crumb = CloudFilesCrumb;

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

	/** The template's view of this element: its state and bound actions. */
	private viewModel(): CloudFilesViewModel {
		return {
			text: this.text,
			providers: this.providers,
			locale: this.locale,
			view: this.view,
			path: this.path,
			entries: this.entries,
			next: this.next,
			loading: this.loading,
			error: this.error,
			query: this.query,
			opening: this.opening,
			progress: this.progress,
			focusIndex: this.focusIndex,
			accepted: (entry) => this.accepted(entry),
			select: (index, entry) => {
				this.focusIndex = index;
				this.activate(entry);
			},
			show: (view) => this.show(view),
			enter: (index) => this.enter(index),
			loadMore: () => void this.load(true),
			choose: (id) => {
				this.active = id;
			},
			signIn: () => void this.signIn(),
			signOut: () => void this.signOut(),
			search: (event) => this.onSearchInput(event),
			openLink: (event) => {
				event.preventDefault();
				const input = (event.currentTarget as HTMLFormElement).querySelector('input');
				if (input) void this.openLink(input.value);
			},
			cancel: () => this.openAbort?.abort(),
			keydown: (event) => this.onKeydown(event),
		};
	}

	protected override render() {
		const p = this.provider;
		if (!p) return nothing;
		const account = p.account();
		const vm = this.viewModel();
		return account ? renderBrowser(vm, p, account) : renderSignedOut(vm, p);
	}
}

export const defineCloudFiles = definer('office-ui-cloud-files', () => OfficeUiCloudFiles, [
	defineSearchField,
]);
