/**
 * collaboration-writeback.ts: elected-writer (role 'owner') PPTX write-back.
 *
 * Only the session owner persists snapshots, eliminating last-save-wins races.
 * On a debounced trigger it reloads the retained source bytes, overlays the
 * live Y.Doc slides (optionally merging back a binding's separately-stored
 * master/layout template elements via `mergeTemplateElements`, so template
 * edits survive), re-serializes to PPTX bytes, and hands them to
 * `config.onWriteBack`.
 *
 * Framework-agnostic: every binding shares this scheduling policy. Bindings
 * with a retained serializer provide it through `serialize`; other bindings
 * use the source-byte reload path below.
 */
import {
	createWriteBackScheduler as createCoreWriteBackScheduler,
	DEFAULT_WRITE_BACK_DEBOUNCE_MS,
	flushWriteBackOnPageClose,
	writeBackStateOf,
	type CollaborationWriteBackState,
	type WriteBackScheduler as CoreWriteBackScheduler,
} from 'ooxml-core/collab';
import { PptxHandler } from 'ooxml-core/pptx';
import type { PptxElement, PptxHandlerSaveOptions, PptxSlide } from 'ooxml-core/pptx';

import type { CollaborationConfig } from '../types';
import type { YDocLike } from './collaboration-sync';
import { readSlidesFromYDoc } from './collaboration-sync';

export interface WriteBackDeps {
	/** The live Y.Doc, or null when disconnected. */
	getYDoc: () => YDocLike | null;
	/** Reuse a binding's retained serializer, with cancellation checks between awaits. */
	serialize?: (isCurrent: () => boolean) => Promise<Uint8Array | null> | Uint8Array | null;
	/** The retained source PPTX bytes to reload before overlaying Y.Doc slides. */
	getSourceBytes?: () => Uint8Array | null;
	/** The per-slide master/layout template element store to merge back. */
	getTemplateElements?: () => Record<string, PptxElement[]>;
	/**
	 * Merge template (master/layout) elements back into the broadcast
	 * (template-free) slides before saving. Defaults to a passthrough that
	 * ignores `getTemplateElements` when omitted, for bindings without
	 * template-mode editing.
	 */
	mergeTemplateElements?: (
		slides: PptxSlide[],
		templateElements: Record<string, PptxElement[]>,
	) => PptxSlide[];
	/**
	 * Session-level save options (view properties, table styles, tags, deck
	 * properties, ...) to persist alongside the Y.Doc slides, built the same
	 * way as the binding's own Save/Export path (`buildDeckSaveOptions` in
	 * `render/deck-save-options`). Without this the write-back reloaded the
	 * source bytes and called `handler.save(merged)` with NO options, so an
	 * owner's write-back file silently dropped every session-level edit that
	 * lives outside `slides` (table style edits, view toggles, tags, deck
	 * properties, ...) even though those edits were visible on screen and
	 * would have been saved by a manual Ctrl+S. Omitted for a binding that has
	 * not wired it yet, which reproduces the pre-existing (options-less) save.
	 */
	getSaveOptions?: () => PptxHandlerSaveOptions | undefined;
	/** Called whenever the write-back status changed; read it with `getState()`. */
	onWriteBackStatus?: () => void;
}

export type WriteBackConfig = Pick<
	CollaborationConfig,
	'role' | 'onWriteBack' | 'writeBackDebounceMs' | 'writeBackRetry' | 'externalSession'
>;

export interface WriteBackScheduler {
	/** Debounce a write-back for the given session (no-op unless role 'owner'). */
	schedule: (config: WriteBackConfig) => void;
	/** Cancel any pending write-back. */
	cancel: () => void;
	/** Write a pending change now; resolves once the host acknowledged it. */
	flush: () => Promise<void>;
	/** The status for the collaboration state; undefined until a write-back was scheduled. */
	getState: () => CollaborationWriteBackState | undefined;
}

/**
 * The PowerPoint write-back: the shared `createWriteBackScheduler` of `ooxml-core/collab` (debounce,
 * one write in flight, promise acknowledgement, retry, status) with PowerPoint's serialisation,
 * the owner and session-readiness gates, and a flush on `pagehide` and when the page is hidden
 * (registered while a change is waiting; browsers do not wait for async work on unload).
 */
export function createWriteBackScheduler(deps: WriteBackDeps): WriteBackScheduler {
	let core: CoreWriteBackScheduler | null = null;
	let coreKey = '';
	let current: WriteBackConfig | null = null;
	let unsubscribe: (() => void) | null = null;
	let detachPageClose: (() => void) | null = null;

	const isReady = (): boolean => current?.externalSession?.getSnapshot().synced ?? true;
	const detachSession = (): void => {
		unsubscribe?.();
		unsubscribe = null;
	};

	function serialize(isCurrentWrite: () => boolean): Promise<Uint8Array | null> {
		const ydoc = deps.getYDoc();
		const isCurrent = (): boolean => isCurrentWrite() && isReady() && deps.getYDoc() === ydoc;
		if (!ydoc || !current?.onWriteBack || !isReady()) return Promise.resolve(null);
		return serializeDoc(ydoc, isCurrent);
	}

	async function serializeDoc(
		ydoc: YDocLike,
		isCurrent: () => boolean,
	): Promise<Uint8Array | null> {
		if (deps.serialize) {
			try {
				const bytes = await deps.serialize(isCurrent);
				return bytes && isCurrent() ? bytes : null;
			} catch (error) {
				// A serializer that gave up because the session changed is not a failed save.
				if (!isCurrent()) return null;
				throw error;
			}
		}
		const sourceBytes = deps.getSourceBytes?.();
		if (!sourceBytes) return null;
		const handler = new PptxHandler();
		await handler.load(sourceBytes.buffer as ArrayBuffer);
		if (!isCurrent()) return null;
		const slides = readSlidesFromYDoc(ydoc);
		const merged = deps.mergeTemplateElements
			? deps.mergeTemplateElements(slides, deps.getTemplateElements?.() ?? {})
			: slides;
		const bytes = await handler.save(merged, deps.getSaveOptions?.());
		return isCurrent() ? bytes : null;
	}

	function coreFor(config: WriteBackConfig): CoreWriteBackScheduler {
		const retry = config.writeBackRetry;
		const key = [
			config.writeBackDebounceMs,
			retry?.attempts,
			retry?.baseDelayMs,
			retry?.maxDelayMs,
		].join('|');
		if (core && key === coreKey) return core;
		core?.dispose();
		coreKey = key;
		core = createCoreWriteBackScheduler({
			debounceMs: config.writeBackDebounceMs ?? DEFAULT_WRITE_BACK_DEBOUNCE_MS,
			...(retry ? { retry } : {}),
			serialize: ({ isCurrent }) => serialize(isCurrent),
			// Read at write time, so a host callback changed meanwhile is the one called.
			write: (bytes) => current?.onWriteBack?.(bytes),
			onStatus: (state) => {
				if (state.status === 'saved' || state.status === 'failed' || state.status === 'idle') {
					detachSession();
				}
				deps.onWriteBackStatus?.();
			},
		});
		return core;
	}

	/**
	 * Stopping or suspending the session: an owner's unsaved change is written once more first
	 * (as Visio does on leave), unless the session already lost its readiness, so a half-synced
	 * document is never written. The session's document is read as it was when cancel was called.
	 */
	function cancel(): void {
		const config = current;
		const ydoc = deps.getYDoc();
		const final =
			core?.dirty && ydoc && config?.onWriteBack && config.role === 'owner' && isReady()
				? { ydoc, onWriteBack: config.onWriteBack }
				: null;
		if (final) {
			void serializeDoc(final.ydoc, () => true)
				.then((bytes) => (bytes ? final.onWriteBack(bytes) : undefined))
				.catch(() => {});
		}
		core?.cancel();
		detachSession();
		detachPageClose?.();
		detachPageClose = null;
	}

	function schedule(config: WriteBackConfig): void {
		if (!config.onWriteBack || config.role !== 'owner' || !deps.getYDoc()) {
			return;
		}
		current = config;
		const scheduler = coreFor(config);
		detachSession();
		if (!isReady()) {
			scheduler.cancel();
			return;
		}
		if (config.externalSession) {
			const detach = config.externalSession.subscribe(() => {
				if (!isReady()) scheduler.cancel();
			});
			// A host may synchronously notify during subscription.
			if (!isReady()) {
				detach();
				return;
			}
			unsubscribe = detach;
		}
		detachPageClose ??= flushWriteBackOnPageClose(scheduler);
		scheduler.schedule();
	}

	return {
		schedule,
		cancel,
		flush: () => core?.flush() ?? Promise.resolve(),
		getState: () => (core ? writeBackStateOf(core.state) : undefined),
	};
}
