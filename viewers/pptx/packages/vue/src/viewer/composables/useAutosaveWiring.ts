import { resolveAutosaveActivation, resolveAutosaveIntervalMs } from 'ooxml-ui/pptx';
import type { AutosaveDisabledReason, DocumentStorage, InlineTextDraft } from 'ooxml-ui/pptx';
/**
 * useAutosaveWiring: the title-bar AutoSave toggle plus the debounced save
 * engine behind it.
 *
 * The gate itself is the shared `resolveAutosaveActivation`: the host's
 * `autosave` prop is a POLICY CEILING and the toggle is the user's PREFERENCE
 * inside it, identically in all five bindings. The title bar has to explain
 * WHICH condition is missing, so the verdict and its reason come back together
 * rather than being recomputed at two call sites.
 */
import type { PptxSlide } from 'pptx-viewer-core';
import type { ComputedRef, Ref, ShallowRef } from 'vue';
import { computed, ref, watch } from 'vue';

import type { UseAutosaveResult } from './useAutosave';
import { useAutosave } from './useAutosave';

export interface UseAutosaveWiringOptions {
	slides: ShallowRef<PptxSlide[]>;
	/**
	 * The per-slide master/layout (template) element store. Edit-template mode
	 * rebuilds this and NOT `slides`, so without it those edits are never
	 * autosaved. See `UseAutosaveOptions.templateElements`.
	 */
	templateElements?: Ref<unknown>;
	/** True while a Design gallery hover preview shows over the slides (see `useAutosave`). */
	paused?: Ref<boolean>;
	/** True while the load pipeline is running; used to clear the dirty flag after a load. */
	loading: Ref<boolean>;
	canEdit: () => boolean;
	/**
	 * The host's `autosave` prop, read as a getter so a mid-session change
	 * applies. `undefined` means the host stated no policy, which permits
	 * autosave; only an explicit `false` vetoes it.
	 */
	autosaveEnabledByHost: () => boolean | undefined;
	/** The host's `autosaveIntervalMs` prop, if any. */
	intervalMs: () => number | undefined;
	/**
	 * File > Options > Save > "Save AutoRecover information every N minutes", in
	 * seconds. Used whenever the host did not state a cadence of its own; Vue
	 * used to ignore it entirely and sit on a 2s debounce forever.
	 */
	optionsIntervalSeconds?: () => number | undefined;
	/** Label the recovery snapshot is stored under (file path, else file name). */
	snapshotName: () => string;
	/** Host `storage` prop: snapshots go to its recovery methods when it has them. */
	storage?: () => DocumentStorage | undefined;
	/**
	 * Serialise the current deck for crash recovery. Must be the plaintext
	 * (`recovery-snapshot`) serialisation, NOT `getContent`: an encrypted
	 * snapshot cannot be reopened, because no recovery path has the password.
	 */
	getRecoverySnapshot: (draft?: InlineTextDraft | null) => Promise<Uint8Array>;
	/** The element whose inline editor is open (its live draft is in the snapshot already). */
	getInlineEditingId?: () => string | null;
	/** Emit the serialised bytes to the host. */
	emitAutosave: (bytes: Uint8Array) => void;
	/** Snapshot a restorable version (version history). */
	captureVersion: (label: string, at: number) => void;
}

export interface UseAutosaveWiringResult {
	autosave: UseAutosaveResult;
	/** The title bar's user-facing AutoSave toggle (defaults on). */
	autosaveEnabled: Ref<boolean>;
	/** Whether autosave is actually running (toggle AND host policy AND gates). */
	autosaveActive: ComputedRef<boolean>;
	toggleAutosave: () => void;
	/** Why autosave is inactive, for the title bar's status message. */
	autosaveDisabledReason: ComputedRef<AutosaveDisabledReason | undefined>;
	/** False when the host forbade autosave: the title-bar switch renders inert. */
	autosaveToggleAvailable: ComputedRef<boolean>;
}

export function useAutosaveWiring(options: UseAutosaveWiringOptions): UseAutosaveWiringResult {
	const autosaveEnabled = ref(true);
	const activation = computed(() =>
		resolveAutosaveActivation({
			hostAutosave: options.autosaveEnabledByHost(),
			userEnabled: autosaveEnabled.value,
			canEdit: options.canEdit(),
			filePath: options.snapshotName(),
		}),
	);
	const autosaveActive = computed(() => activation.value.active);
	const autosaveDisabledReason = computed(() => activation.value.reason);
	const autosaveToggleAvailable = computed(() => activation.value.toggleAvailable);

	function toggleAutosave(): void {
		// Inert when the host passed `autosave={false}`: a preference cannot
		// exceed the policy, so the switch must not move.
		if (activation.value.toggleAvailable) {
			autosaveEnabled.value = !autosaveEnabled.value;
		}
	}

	const autosave = useAutosave({
		slides: options.slides,
		...(options.templateElements ? { templateElements: options.templateElements } : {}),
		...(options.paused ? { paused: options.paused } : {}),
		enabled: autosaveActive,
		// A computed, not a snapshot: the engine re-reads it each time the timer
		// is armed, so a host changing the AutoRecover cadence (File > Options >
		// Save) takes effect without a remount.
		intervalMs: computed(() =>
			resolveAutosaveIntervalMs({
				hostIntervalMs: options.intervalMs(),
				optionsIntervalSeconds: options.optionsIntervalSeconds?.(),
			}),
		),
		filePath: options.snapshotName,
		...(options.storage ? { storage: options.storage } : {}),
		// The shared recovery store (IndexedDB, or the host's recovery methods).
		// Recovery reads these bytes back with no password, so they are
		// deliberately a plain ZIP even when the deck is protected.
		// The open inline editor's rich draft, else the plain autosave draft.
		serialize: (draft) => options.getRecoverySnapshot(draft),
		...(options.getInlineEditingId ? { getInlineEditingId: options.getInlineEditingId } : {}),
		onSaved: (bytes) => {
			options.emitAutosave(bytes);
			options.captureVersion('Autosave', Date.now());
		},
	});

	// Loading a deck reassigns `slides`, which the autosave watcher counts as an
	// edit; clear it once loading settles so a freshly opened deck reads "Saved
	// to this PC" in the title bar, matching React.
	watch(options.loading, (now, was) => {
		if (was && !now) {
			autosave.isDirty.value = false;
		}
	});

	return {
		autosave,
		autosaveEnabled,
		autosaveActive,
		toggleAutosave,
		autosaveDisabledReason,
		autosaveToggleAvailable,
	};
}
