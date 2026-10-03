import {
	DocumentHistory,
	EMPTY_EDIT_STATE,
	type ViewerEditState,
	type VsdxExportResult,
} from './document-history.js';
import { createWorkerEditor, snapshotEdits, type CancellableEditor } from './worker-editor.js';
import { MAX_INPUT_BYTES } from './scene-validation.js';
import { parseVsdx, type VisioDocument, type VisioEdit } from 'ooxml-core/visio';
import {
	EMPTY_LAYER_OVERRIDES,
	documentVisibility,
	visibleSelection,
	type LayerVisibilityOverride,
} from './viewer-layers.js';
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
	readonly edit: ViewerEditState;
	readonly pageIndex: number;
	readonly zoom: number;
	readonly loading: boolean;
	readonly error: Error | null;
	readonly search: TextSearchState;
	/** Viewer-only display overrides, independent of saved model and print policy. */
	readonly layerVisibilityOverrides: readonly LayerVisibilityOverride[];
	readonly selectedShape: { id: string; name: string; pageId?: string } | null;
}
type Parser = CancellableParser;
type EventListener = <K extends keyof ViewerEvents>(name: K, detail: ViewerEvents[K]) => void;
export class ViewerController {
	#state: ViewerState = Object.freeze({
		document: null,
		edit: EMPTY_EDIT_STATE,
		pageIndex: 0,
		zoom: 1,
		loading: false,
		error: null,
		selectedShape: null,
		search: EMPTY_TEXT_SEARCH,
		layerVisibilityOverrides: EMPTY_LAYER_OVERRIDES,
	});
	#subscribers = new Set<(state: ViewerState) => void>();
	#events = new Set<EventListener>();
	#loadId = 0;
	#destroyed = false;
	#revision = 0;
	#documentGeneration = 0;
	#sourceGeneration = 0;
	#searchIndex: DocumentTextIndex | null = null;
	#history: DocumentHistory | null = null;
	#editId = 0;
	#visible = documentVisibility(null);
	constructor(
		private readonly parser: Parser = parseVsdx,
		private readonly listenerError: (error: unknown) => void = (error) => {
			if (typeof globalThis.reportError === 'function') globalThis.reportError(error);
			else console.error('Visio viewer listener failed:', error);
		},
		private readonly editor: CancellableEditor = createWorkerEditor(),
	) {}
	get state(): ViewerState {
		return this.#state;
	}
	/** Accepted document replacements, including the same object, invalidate captured artifacts. */
	get documentGeneration(): number {
		this.#assertAlive();
		return this.#documentGeneration;
	}
	/** Accepted source/model replacements invalidate drafts; edit/undo/redo keep this source epoch. */
	get sourceGeneration(): number {
		this.#assertAlive();
		return this.#sourceGeneration;
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
		const visible = documentVisibility(document);
		++this.#documentGeneration;
		++this.#sourceGeneration;
		this.#invalidateEdit();
		this.#history = null;
		this.#loadId++;
		this.parser.cancel?.();
		this.#searchIndex = null;
		this.#visible = visible;
		this.#change({
			document,
			edit: EMPTY_EDIT_STATE,
			pageIndex: 0,
			loading: false,
			error: null,
			selectedShape: null,
			search: EMPTY_TEXT_SEARCH,
			layerVisibilityOverrides: EMPTY_LAYER_OVERRIDES,
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
		if (
			shape &&
			!visibleSelection(this.#state.document, this.#state.pageIndex, shape, this.#visible)
		)
			return;
		if (this.#change({ selectedShape: shape, search: this.#inactiveSearch() }))
			this.#emit('shape-select', shape);
	}
	/** Override one source-page layer for viewing. null restores its saved display flag. */
	setLayerVisibility(pageId: string, layerId: string, visible: boolean | null): void {
		this.#assertAlive();
		const page = this.#state.document?.pages.find((candidate) => candidate.id === pageId);
		if (!page) throw new Error('The layer override page does not belong to this document.');
		if (!page.layers?.some((layer) => layer.id === layerId))
			throw new Error('The layer does not belong to this page.');
		if (visible !== null && typeof visible !== 'boolean')
			throw new Error('Layer visibility must be boolean or null.');
		const previous = this.#state.layerVisibilityOverrides;
		const found = previous.find((entry) => entry.pageId === pageId && entry.layerId === layerId);
		if (visible === null ? !found : found?.visible === visible) return;
		const next = previous.filter((entry) => entry !== found);
		if (visible !== null) next.push(Object.freeze({ pageId, layerId, visible }));
		this.#setLayerOverrides(Object.freeze(next));
	}
	/** Reset a source page, or all pages when omitted. Does not modify the document. */
	resetLayerVisibility(pageId?: string): void {
		this.#assertAlive();
		if (pageId !== undefined && !this.#state.document?.pages.some((page) => page.id === pageId))
			throw new Error('The layer override page does not belong to this document.');
		const previous = this.#state.layerVisibilityOverrides;
		const next =
			pageId === undefined
				? EMPTY_LAYER_OVERRIDES
				: Object.freeze(previous.filter((entry) => entry.pageId !== pageId));
		if (next.length !== previous.length) this.#setLayerOverrides(next);
	}
	#setLayerOverrides(layerVisibilityOverrides: readonly LayerVisibilityOverride[]): void {
		const visible = documentVisibility(this.#state.document, layerVisibilityOverrides);
		const searchIndex = this.#state.search.query
			? indexDocumentText(this.#state.document, (shape) => visible.get(shape) === true)
			: null;
		const search = searchIndex
			? searchDocumentText(searchIndex, this.#state.search.query)
			: EMPTY_TEXT_SEARCH;
		const cleared =
			!!this.#state.selectedShape &&
			!visibleSelection(
				this.#state.document,
				this.#state.pageIndex,
				this.#state.selectedShape,
				visible,
			);
		this.#visible = visible;
		this.#searchIndex = searchIndex;
		if (
			this.#change({
				layerVisibilityOverrides,
				search,
				...(cleared ? { selectedShape: null } : {}),
			}) &&
			cleared
		)
			this.#emit('shape-select', null);
	}
	/** Search does not change the page or selection until explicit result navigation. */
	setSearchQuery(query: string): void {
		this.#assertAlive();
		validateSearchQuery(query);
		if (query === this.#state.search.query) return;
		this.#searchIndex ??= indexDocumentText(
			this.#state.document,
			(shape) => this.#visible.get(shape) === true,
		);
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
		this.#assertAlive();
		if (bytes.byteLength > MAX_INPUT_BYTES)
			throw new Error('This viewer accepts files up to 32 MiB.');
		const owned =
			bytes instanceof Uint8Array ? Uint8Array.from(bytes) : new Uint8Array(bytes.slice(0));
		return this.loadSource(() => owned);
	}
	/** Reading bytes shares the same request epoch and error state as parsing them. */
	async loadSource(
		read: () => Uint8Array | ArrayBuffer | Promise<Uint8Array | ArrayBuffer>,
	): Promise<void> {
		this.#assertAlive();
		const id = ++this.#loadId;
		this.#invalidateEdit();
		this.parser.cancel?.();
		this.#change({ loading: true, error: null, edit: this.#history?.state ?? EMPTY_EDIT_STATE });
		if (this.#destroyed || id !== this.#loadId) return;
		let document: VisioDocument;
		let history: DocumentHistory;
		let visible: ReturnType<typeof documentVisibility>;
		try {
			const bytes = await read();
			if (this.#destroyed || id !== this.#loadId) return;
			if (bytes.byteLength > MAX_INPUT_BYTES)
				throw new Error('This viewer accepts files up to 32 MiB.');
			const owned =
				bytes instanceof Uint8Array ? Uint8Array.from(bytes) : new Uint8Array(bytes.slice(0));
			history = new DocumentHistory(owned);
			document = await this.parser(Uint8Array.from(owned));
			if (this.#destroyed || id !== this.#loadId) return;
			assertViewableDocument(document);
			visible = documentVisibility(document);
		} catch (cause) {
			if (this.#destroyed || id !== this.#loadId) return;
			const error = cause instanceof Error ? cause : new Error(String(cause));
			this.#change({ loading: false, error });
			if (!this.#destroyed && id === this.#loadId) this.#emit('document-error', error);
			throw error;
		}
		this.#history = history;
		this.#searchIndex = null;
		this.#visible = visible;
		++this.#documentGeneration;
		++this.#sourceGeneration;
		this.#change({
			document,
			edit: history.state,
			pageIndex: 0,
			loading: false,
			error: null,
			selectedShape: null,
			search: EMPTY_TEXT_SEARCH,
			layerVisibilityOverrides: EMPTY_LAYER_OVERRIDES,
		});
		if (!this.#destroyed && id === this.#loadId) this.#emit('document-load', document);
	}
	cancelLoad(): void {
		this.#assertAlive();
		++this.#loadId;
		this.parser.cancel?.();
		this.#invalidateEdit();
		this.#change({ loading: false, edit: this.#history?.state ?? EMPTY_EDIT_STATE });
	}
	destroy(): void {
		if (this.#destroyed) return;
		this.#destroyed = true;
		this.#invalidateEdit();
		this.#history = null;
		++this.#documentGeneration;
		++this.#sourceGeneration;
		++this.#loadId;
		this.parser.cancel?.();
		this.#subscribers.clear();
		this.#events.clear();
		this.#searchIndex = null;
		this.#visible = documentVisibility(null);
		this.#state = Object.freeze({
			document: null,
			edit: EMPTY_EDIT_STATE,
			pageIndex: 0,
			zoom: 1,
			loading: false,
			error: null,
			selectedShape: null,
			search: EMPTY_TEXT_SEARCH,
			layerVisibilityOverrides: EMPTY_LAYER_OVERRIDES,
		});
	}
	/** Replace one source-backed local plain-text target. Core decides target support. */
	async replacePlainText(pageId: string, shapeId: string, text: string): Promise<void> {
		return this.applyEdits([{ type: 'replace-plain-text', pageId, shapeId, text }]);
	}
	/** Atomic source-backed edits. Core owns protection, dependency and target validation. */
	async applyEdits(edits: readonly VisioEdit[]): Promise<void> {
		return this.#mutate('edit', snapshotEdits(edits));
	}
	async undo(): Promise<void> {
		return this.#mutate('undo');
	}
	async redo(): Promise<void> {
		return this.#mutate('redo');
	}
	cancelEdit(): void {
		this.#assertAlive();
		if (!this.#state.edit.busy) return;
		this.#invalidateEdit();
		this.parser.cancel?.();
		this.#change({ edit: this.#history?.state ?? EMPTY_EDIT_STATE });
	}
	exportVsdx(): VsdxExportResult {
		this.#assertAlive();
		if (!this.#history)
			throw new Error('Load a VSDX file before downloading a source-backed copy.');
		if (this.#state.edit.busy || this.#state.loading)
			throw new Error('Wait for the current document operation before downloading.');
		return this.#history.export();
	}
	#invalidateEdit(): void {
		++this.#editId;
		this.editor.cancel?.();
	}
	async #mutate(kind: 'edit' | 'undo' | 'redo', commands?: readonly VisioEdit[]): Promise<void> {
		this.#assertAlive();
		const history = this.#history;
		if (!history)
			throw new Error('Load a VSDX file before editing. Model-only documents are read only.');
		if (this.#state.loading || this.#state.edit.busy)
			throw new Error('Another document operation is in progress.');
		const target =
			kind === 'undo' ? history.undoTarget : kind === 'redo' ? history.redoTarget : undefined;
		if (kind !== 'edit' && !target) return;
		const id = ++this.#editId;
		const loadId = this.#loadId;
		const current = () =>
			!this.#destroyed &&
			id === this.#editId &&
			loadId === this.#loadId &&
			history === this.#history;
		const assertCurrent = () => {
			if (!current())
				throw new DOMException('The diagram edit was superseded or cancelled.', 'AbortError');
		};
		this.#change({ edit: Object.freeze({ ...history.state, busy: true }) });
		assertCurrent();
		try {
			let document: VisioDocument;
			let edited: Awaited<ReturnType<CancellableEditor>> | undefined;
			if (kind === 'edit') {
				edited = await this.editor(Uint8Array.from(history.current.bytes), commands!);
				assertCurrent();
				if (edited.bytes.byteLength > MAX_INPUT_BYTES)
					throw new Error('The modified VSDX exceeds 32 MiB.');
				document = edited.document;
			} else document = await this.parser(Uint8Array.from(target!.bytes));
			assertCurrent();
			assertViewableDocument(document);
			if (edited && !edited.changedParts.length) {
				this.#change({ edit: history.state });
				return;
			}
			const visible = documentVisibility(document, this.#state.layerVisibilityOverrides);
			const searchIndex = this.#state.search.query
				? indexDocumentText(document, (shape) => visible.get(shape) === true)
				: null;
			const search = searchIndex
				? searchDocumentText(searchIndex, this.#state.search.query)
				: EMPTY_TEXT_SEARCH;
			const selection = this.#state.selectedShape;
			const selectedShape =
				selection && visibleSelection(document, this.#state.pageIndex, selection, visible)
					? selection
					: null;
			// No external callbacks occur between history acceptance and model acceptance.
			if (edited) history.append(edited.bytes, edited.diagnostics);
			else history.move(target!);
			this.#visible = visible;
			this.#searchIndex = searchIndex;
			++this.#documentGeneration;
			if (
				this.#change({ document, edit: history.state, search, selectedShape, error: null }) &&
				current()
			)
				this.#emit('document-change', { document, dirty: history.state.dirty, kind });
		} catch (cause) {
			if (!current())
				throw new DOMException('The diagram edit was superseded or cancelled.', 'AbortError');
			const error = cause instanceof Error ? cause : new Error(String(cause));
			this.#change({ edit: Object.freeze({ ...history.state, error }) });
			throw error;
		}
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
