/**
 * useAutosaveRecovery: the Vue adapter of the shared crash-recovery offer
 * (`createAutosaveRecoveryController` in `ooxml-ui/pptx`). The decision, the
 * copy and the actions are shared; this composable only owns the reactivity.
 */
import {
	createAutosaveRecoveryController,
	discardAutosaveRecovery,
	probeAutosaveRecovery,
} from 'ooxml-ui/pptx';
import type { AutosaveRecord, AutosaveRecoveryPrompt, DocumentStorage } from 'ooxml-ui/pptx';
import { computed, onScopeDispose, watch } from 'vue';
import type { Ref } from 'vue';

import { useController } from './useController';

export interface UseAutosaveRecoveryOptions {
	/** IndexedDB key of the open deck (file path, else file name). */
	filePath: () => string | undefined;
	/** Public document name rendered in the recovery message. */
	fileName: () => string | undefined;
	/** Host `storage` prop: the snapshot is read from its recovery methods when it has them. */
	storage?: () => DocumentStorage | undefined;
	/** True while the load pipeline is running. */
	loading: Ref<boolean>;
	/** Load error, if any. */
	error: Ref<string | null | undefined>;
	/** Slides currently rendered. */
	slideCount: () => number;
	/**
	 * Whether the host permits autosave at all (`autosave` prop not `false`). A
	 * user who merely switched the toggle off is still offered a pre-crash
	 * snapshot; a host that forbade the feature is not.
	 */
	autosaveAllowed: () => boolean;
	/** The host's `autosave-recovery` event: called with the record when a snapshot is offered. */
	onOffered?: (record: AutosaveRecord) => void;
	/** Load the recovered bytes into the viewer. */
	onRestore: (bytes: Uint8Array) => void;
}

export interface UseAutosaveRecoveryResult {
	/** What the dialog should say, or null when there is nothing to offer. */
	prompt: Ref<AutosaveRecoveryPrompt | null>;
	discarding: Ref<boolean>;
	restore: () => void;
	discard: () => Promise<void>;
}

export function useAutosaveRecovery(
	options: UseAutosaveRecoveryOptions,
): UseAutosaveRecoveryResult {
	const controller = createAutosaveRecoveryController({
		getStorage: () => options.storage?.(),
		load: (bytes) => options.onRestore(bytes),
		onOffered: (record) => options.onOffered?.(record),
		probe: probeAutosaveRecovery,
		discardRecord: discardAutosaveRecovery,
	});
	const state = useController(controller);
	onScopeDispose(() => controller.destroy());

	watch(
		() => [options.loading.value, options.error.value, options.slideCount(), options.filePath()],
		() =>
			controller.check({
				filePath: options.filePath(),
				fileName: options.fileName(),
				loading: options.loading.value,
				error: options.error.value ?? null,
				slideCount: options.slideCount(),
				autosaveAllowed: options.autosaveAllowed(),
			}),
		{ immediate: true },
	);

	return {
		prompt: computed(() => state.value.prompt) as Ref<AutosaveRecoveryPrompt | null>,
		discarding: computed(() => state.value.discarding) as Ref<boolean>,
		restore: () => {
			void controller.restore().catch(() => {
				// The load path reports its own error.
			});
		},
		discard: () => controller.discard(),
	};
}
