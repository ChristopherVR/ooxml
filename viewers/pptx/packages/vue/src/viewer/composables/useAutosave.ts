import { createAutosaveController, saveAutosaveSnapshot } from 'ooxml-ui/pptx';
import type { DocumentStorage, PptxAutosaveStatusKind } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import { computed, onScopeDispose, toValue, watch } from 'vue';
import type { Ref, WritableComputedRef } from 'vue';

import { useController } from './useController';

/**
 * Autosave status, surfaced for status-pill rendering: the `state` of the one
 * shared `PptxAutosaveStatus` (`idle`, `disabled`, `saving`, `saved`, `error`).
 */
export type AutosaveStatus = PptxAutosaveStatusKind;

/**
 * Timer-injection seam so the interval is testable with `vi.useFakeTimers()`.
 * Defaults to the global `setTimeout` / `clearTimeout`.
 */
export interface AutosaveTimerApi {
	setTimer: (handler: () => void, timeoutMs: number) => number;
	clearTimer: (id: number) => void;
}

export interface UseAutosaveOptions {
	/**
	 * The editor's reactive slide array. Because the editor reassigns this
	 * array immutably on every edit, a shallow `watch(slides, …)` fires per
	 * edit, no `deep` needed.
	 */
	slides: Ref<PptxSlide[]>;
	/**
	 * The separate per-slide store of master/layout (template) elements. An
	 * edit made in edit-template mode rebuilds ONLY this map, so it is watched
	 * too; otherwise a user editing a master or layout got no crash recovery.
	 */
	templateElements?: Ref<unknown>;
	/**
	 * True while a Design gallery hover preview shows over the slides: the stores change but the
	 * document does not, so nothing is marked dirty and no timer is armed.
	 */
	paused?: Ref<boolean>;
	/** Master on/off switch: the resolved shared activation, not the raw toggle. */
	enabled?: Ref<boolean> | boolean;
	/**
	 * The AutoRecover interval in milliseconds. A `Ref` is re-read each time the
	 * timer is armed, and a pending timer is re-armed when it changes.
	 */
	intervalMs: number | Ref<number>;
	/**
	 * Serialise the deck for recovery (plaintext). With it, the engine writes the
	 * snapshot itself under {@link filePath}; without it, {@link onSave} owns the
	 * whole save.
	 */
	serialize?: () => Promise<Uint8Array | null>;
	/** Snapshot key; autosave is inert while it is empty (used with {@link serialize}). */
	filePath?: () => string | undefined;
	/** Host `storage`: snapshots go to its recovery methods when it has them. */
	storage?: () => DocumentStorage | undefined;
	/** Called with the bytes of each snapshot written (used with {@link serialize}). */
	onSaved?: (bytes: Uint8Array) => void;
	/**
	 * Host save callback for callers that persist the bytes themselves. May be
	 * sync or async; rejections flip status to `'error'`.
	 */
	onSave?: () => void | Promise<void>;
	/** Injectable timer API for testing. Defaults to the global timers. */
	timers?: AutosaveTimerApi;
}

export interface UseAutosaveResult {
	/** Current autosave lifecycle status. */
	status: Ref<AutosaveStatus>;
	/** Epoch ms of the last successful save, or `null`. */
	lastSavedAt: Ref<number | null>;
	/** Force an immediate save, bypassing the interval. */
	saveNow: () => Promise<void>;
	/**
	 * True when there are edits no snapshot has captured yet. Setting it true is
	 * an edit (it starts the interval); setting it false clears them.
	 */
	isDirty: WritableComputedRef<boolean>;
}

const NO_BYTES = new Uint8Array(0);

/**
 * `useAutosave`: the Vue adapter of the shared recovery autosave engine
 * (`createAutosaveController` in `ooxml-ui/pptx`). Each reassignment of the
 * watched stores is an edit; the engine writes one snapshot an AutoRecover
 * interval after the first unsaved edit. The timer is torn down on scope dispose.
 */
export function useAutosave(options: UseAutosaveOptions): UseAutosaveResult {
	const { slides, templateElements, enabled = true, intervalMs, paused } = options;
	const timers = options.timers;
	const ownsWrite = Boolean(options.serialize);

	const controller = createAutosaveController({
		getActivation: () => ({ active: toValue(enabled) !== false }),
		getFilePath: () => (ownsWrite ? options.filePath?.() : 'autosave'),
		getIntervalMs: () => toValue(intervalMs),
		serialize: async () => {
			if (options.serialize) {
				return options.serialize();
			}
			await options.onSave?.();
			return NO_BYTES;
		},
		getStorage: () => options.storage?.(),
		onSaved: (bytes) => {
			if (ownsWrite) {
				options.onSaved?.(bytes);
			}
		},
		writeSnapshot: ownsWrite ? saveAutosaveSnapshot : async () => undefined,
		...(timers
			? {
					timers: {
						setTimeout: (handler, ms) => timers.setTimer(handler, ms),
						clearTimeout: (id) => timers.clearTimer(id as number),
					},
				}
			: {}),
	});
	const state = useController(controller);

	const status = computed(() => state.value.status.state);
	const lastSavedAt = computed(() => state.value.lastSavedAt);
	const isDirty = computed({
		get: () => state.value.dirty,
		set: (next: boolean) => {
			if (next) {
				controller.markEdited();
			} else {
				controller.markClean();
			}
		},
	});

	// Both stores are watched because an edit lands in exactly one of them: a
	// normal edit rebuilds `slides`, a template-mode edit rebuilds the template
	// map. Flushed synchronously so an edit marks the document dirty at once.
	watch(
		templateElements ? [slides, templateElements] : [slides],
		() => {
			if (!paused?.value) {
				controller.markEdited();
			}
		},
		{ flush: 'sync' },
	);

	// The activation and the cadence (File > Options > Save) apply at once.
	watch(
		() => [toValue(enabled), toValue(intervalMs), options.filePath?.()],
		() => controller.refresh(),
	);

	onScopeDispose(() => controller.destroy());

	return {
		status: status as Ref<AutosaveStatus>,
		lastSavedAt: lastSavedAt as Ref<number | null>,
		saveNow: () => controller.saveNow(),
		isDirty,
	};
}
