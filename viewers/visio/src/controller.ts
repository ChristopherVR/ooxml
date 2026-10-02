import { parseVsdx, type VisioDocument } from 'ooxml-core/visio';
import { assertViewableDocument } from './scene-validation.js';
import type { CancellableParser } from './worker-parser.js';
import type { ViewerEvents } from './contract.js';
import {
	EMPTY_TEXT_SEARCH,
	indexDocumentText,
	searchDocumentText,
	validateSearchQuery,
	type DocumentTextIndex,
	type TextSearchState,
} from './document-text-search.js';

export interface ViewerState {
	readonly document: VisioDocument | null;
	readonly pageIndex: number;
	readonly zoom: number;
	readonly loading: boolean;
	readonly error: Error | null;
	readonly search: TextSearchState;
	readonly selectedShape: { id: string; name: string; pageId?: string } | null;
}
type Parser = CancellableParser;
type EventListener = <K extends keyof ViewerEvents>(name: K, detail: ViewerEvents[K]) => void;
export class ViewerController {
	#state: ViewerState = Object.freeze({
		document: null,
		pageIndex: 0,
		zoom: 1,
		loading: false,
		error: null,
		selectedShape: null,
		search: EMPTY_TEXT_SEARCH,
	});
	#subscribers = new Set<(state: ViewerState) => void>();
	#events = new Set<EventListener>();
	#loadId = 0;
	#destroyed = false;
	#revision = 0;
	#documentGeneration = 0;
	#searchIndex: DocumentTextIndex | null = null;
	constructor(
		private readonly parser: Parser = parseVsdx,
		private readonly listenerError: (error: unknown) => void = (error) => {
			if (typeof globalThis.reportError === 'function') globalThis.reportError(error);
			else console.error('Visio viewer listener failed:', error);
		},
	) {}
	get state(): ViewerState {
		return this.#state;
	}
	/** Accepted document replacements, including the same object, invalidate captured artifacts. */
	get documentGeneration(): number {
		this.#assertAlive();
		return this.#documentGeneration;
	}
	subscribe(listener: (state: ViewerState) => void): () => void {
		this.#assertAlive();
		this.#subscribers.add(listener);
		if (!this.#notify(() => listener(this.#state))) this.#subscribers.delete(listener);
		return () => this.#subscribers.delete(listener);
	}
	onEvent(listener: EventListener): () => void {
		this.#assertAlive();
		this.#events.add(listener);
		return () => this.#events.delete(listener);
	}
	setDocument(document: VisioDocument | null): void {
		this.#assertAlive();
		if (document) assertViewableDocument(document);
		++this.#documentGeneration;
		this.#loadId++;
		this.parser.cancel?.();
		this.#searchIndex = null;
		this.#change({
			document,
			pageIndex: 0,
			loading: false,
			error: null,
			selectedShape: null,
			search: EMPTY_TEXT_SEARCH,
		});
	}
	setPage(index: number): void {
		this.#assertAlive();
		const count = this.#state.document?.pages.length ?? 0;
		const next = Math.max(0, Math.min(count - 1, Number.isFinite(index) ? Math.trunc(index) : 0));
		if (next === this.#state.pageIndex) return;
		if (this.#change({ pageIndex: next, selectedShape: null, search: this.#inactiveSearch() }))
			this.#emit('page-change', next);
	}
	setZoom(zoom: number): void {
		this.#assertAlive();
		const next = Math.max(0.1, Math.min(8, Number.isFinite(zoom) ? zoom : 1));
		if (next === this.#state.zoom) return;
		if (this.#change({ zoom: next })) this.#emit('zoom-change', next);
	}
	selectShape(shape: ViewerState['selectedShape']): void {
		this.#assertAlive();
		if (this.#change({ selectedShape: shape, search: this.#inactiveSearch() }))
			this.#emit('shape-select', shape);
	}
	/** Search does not change the page or selection until explicit result navigation. */
	setSearchQuery(query: string): void {
		this.#assertAlive();
		validateSearchQuery(query);
		if (query === this.#state.search.query) return;
		this.#searchIndex ??= indexDocumentText(this.#state.document);
		this.#change({ search: searchDocumentText(this.#searchIndex, query) });
	}
	selectSearchResult(index: number): void {
		this.#assertAlive();
		if (!Number.isSafeInteger(index)) return;
		const result = this.#state.search.results[index];
		if (!result) return;
		const pageChanged = this.#state.pageIndex !== result.pageIndex;
		if (
			!this.#change({
				search: Object.freeze({ ...this.#state.search, activeIndex: index }),
				pageIndex: result.pageIndex,
				...(pageChanged ? { selectedShape: null } : {}),
			})
		)
			return;
		const revision = this.#revision;
		if (pageChanged) this.#emit('page-change', result.pageIndex);
		// Page callbacks/subscribers may replace the document or issue newer navigation.
		if (this.#destroyed || revision !== this.#revision) return;
		const shape = { id: result.shapeId, name: result.shapeName, pageId: result.pageId };
		if (this.#change({ selectedShape: shape })) this.#emit('shape-select', shape);
	}
	nextSearchResult(): void {
		this.#assertAlive();
		const { activeIndex, results } = this.#state.search;
		if (results.length) this.selectSearchResult((activeIndex + 1) % results.length);
	}
	previousSearchResult(): void {
		this.#assertAlive();
		const { activeIndex, results } = this.#state.search;
		if (results.length)
			this.selectSearchResult(activeIndex <= 0 ? results.length - 1 : activeIndex - 1);
	}
	#inactiveSearch(): TextSearchState {
		return this.#state.search.activeIndex < 0
			? this.#state.search
			: Object.freeze({ ...this.#state.search, activeIndex: -1 });
	}
	async load(bytes: Uint8Array | ArrayBuffer): Promise<void> {
		return this.loadSource(() => bytes);
	}
	/** Reading bytes shares the same request epoch and error state as parsing them. */
	async loadSource(
		read: () => Uint8Array | ArrayBuffer | Promise<Uint8Array | ArrayBuffer>,
	): Promise<void> {
		this.#assertAlive();
		const id = ++this.#loadId;
		this.parser.cancel?.();
		this.#change({ loading: true, error: null });
		if (this.#destroyed || id !== this.#loadId) return;
		let document: VisioDocument;
		try {
			const bytes = await read();
			if (this.#destroyed || id !== this.#loadId) return;
			document = await this.parser(bytes);
			if (this.#destroyed || id !== this.#loadId) return;
			assertViewableDocument(document);
		} catch (cause) {
			if (this.#destroyed || id !== this.#loadId) return;
			const error = cause instanceof Error ? cause : new Error(String(cause));
			this.#change({ loading: false, error });
			if (!this.#destroyed && id === this.#loadId) this.#emit('document-error', error);
			throw error;
		}
		this.#searchIndex = null;
		++this.#documentGeneration;
		this.#change({
			document,
			pageIndex: 0,
			loading: false,
			error: null,
			selectedShape: null,
			search: EMPTY_TEXT_SEARCH,
		});
		if (!this.#destroyed && id === this.#loadId) this.#emit('document-load', document);
	}
	cancelLoad(): void {
		this.#assertAlive();
		++this.#loadId;
		this.parser.cancel?.();
		this.#change({ loading: false });
	}
	destroy(): void {
		if (this.#destroyed) return;
		this.#destroyed = true;
		++this.#documentGeneration;
		++this.#loadId;
		this.parser.cancel?.();
		this.#subscribers.clear();
		this.#events.clear();
		this.#searchIndex = null;
		this.#state = Object.freeze({
			document: null,
			pageIndex: 0,
			zoom: 1,
			loading: false,
			error: null,
			selectedShape: null,
			search: EMPTY_TEXT_SEARCH,
		});
	}
	#assertAlive(): void {
		if (this.#destroyed) throw new Error('The viewer has been destroyed.');
	}
	#change(change: Partial<ViewerState>): boolean {
		const revision = ++this.#revision;
		this.#state = Object.freeze({ ...this.#state, ...change });
		for (const listener of [...this.#subscribers]) {
			if (this.#destroyed || revision !== this.#revision) break;
			if (this.#subscribers.has(listener)) this.#notify(() => listener(this.#state));
		}
		return !this.#destroyed && revision === this.#revision;
	}
	#notify(callback: () => void): boolean {
		try {
			callback();
			return true;
		} catch (error) {
			try {
				this.listenerError(error);
			} catch {
				/* A host error reporter must not corrupt viewer state. */
			}
			return false;
		}
	}
	#emit<K extends keyof ViewerEvents>(name: K, detail: ViewerEvents[K]): void {
		const revision = this.#revision;
		for (const listener of [...this.#events]) {
			if (this.#destroyed || revision !== this.#revision) break;
			if (this.#events.has(listener)) this.#notify(() => listener(name, detail));
		}
	}
}
