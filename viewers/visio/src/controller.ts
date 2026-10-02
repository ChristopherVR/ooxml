import { parseVsdx, type VisioDocument } from 'ooxml-core/visio';
import { assertViewableDocument } from './scene-validation.js';
import type { CancellableParser } from './worker-parser.js';
import type { ViewerEvents } from './contract.js';

export interface ViewerState {
	readonly document: VisioDocument | null;
	readonly pageIndex: number;
	readonly zoom: number;
	readonly loading: boolean;
	readonly error: Error | null;
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
	});
	#subscribers = new Set<(state: ViewerState) => void>();
	#events = new Set<EventListener>();
	#loadId = 0;
	#destroyed = false;
	#revision = 0;
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
		this.#loadId++;
		this.parser.cancel?.();
		this.#change({ document, pageIndex: 0, loading: false, error: null, selectedShape: null });
	}
	setPage(index: number): void {
		this.#assertAlive();
		const count = this.#state.document?.pages.length ?? 0;
		const next = Math.max(0, Math.min(count - 1, Number.isFinite(index) ? Math.trunc(index) : 0));
		if (next === this.#state.pageIndex) return;
		if (this.#change({ pageIndex: next, selectedShape: null })) this.#emit('page-change', next);
	}
	setZoom(zoom: number): void {
		this.#assertAlive();
		const next = Math.max(0.1, Math.min(8, Number.isFinite(zoom) ? zoom : 1));
		if (next === this.#state.zoom) return;
		if (this.#change({ zoom: next })) this.#emit('zoom-change', next);
	}
	selectShape(shape: ViewerState['selectedShape']): void {
		this.#assertAlive();
		if (this.#change({ selectedShape: shape })) this.#emit('shape-select', shape);
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
		this.#change({ document, pageIndex: 0, loading: false, error: null, selectedShape: null });
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
		++this.#loadId;
		this.parser.cancel?.();
		this.#subscribers.clear();
		this.#events.clear();
		this.#state = Object.freeze({
			document: null,
			pageIndex: 0,
			zoom: 1,
			loading: false,
			error: null,
			selectedShape: null,
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
