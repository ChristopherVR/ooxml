/**
 * `useRehearseTimings`: Rehearse Timings for the Vue binding.
 *
 * The clocks, pause bookkeeping, per-slide recording and the display tick live
 * in the shared `createPresentationRehearsalController`; this composable reads
 * it through the one store adapter (`useController`) and names its state for
 * the HUD and summary templates.
 *
 * @module composables/useRehearseTimings
 */

import { createPresentationRehearsalController } from 'ooxml-ui/pptx';
import type { PresentationRehearsalControllerOptions } from 'ooxml-ui/pptx';
import { computed, onScopeDispose, getCurrentScope } from 'vue';
import type { ComputedRef } from 'vue';

import { useController } from './useController';

/**
 * Format a millisecond duration as `M:SS` (minutes are not zero-padded, seconds
 * are), as the React HUD and summary do.
 */
export function formatRehearseMs(ms: number): string {
	const totalSeconds = Math.max(0, Math.floor(ms / 1000));
	const minutes = Math.floor(totalSeconds / 60);
	const seconds = totalSeconds % 60;
	return `${String(minutes)}:${String(seconds).padStart(2, '0')}`;
}

export interface UseRehearseTimingsOptions extends PresentationRehearsalControllerOptions {
	/** Called with the recorded per-slide timings (ms) when the user saves. */
	onSave?: (timings: Record<number, number>) => void;
}

export interface UseRehearseTimingsResult {
	rehearsing: ComputedRef<boolean>;
	paused: ComputedRef<boolean>;
	showSummary: ComputedRef<boolean>;
	/** Recorded per-slide timings (slide index to ms). */
	recordedTimings: ComputedRef<Readonly<Record<number, number>>>;
	slideElapsedMs: ComputedRef<number>;
	totalElapsedMs: ComputedRef<number>;
	totalRecordedMs: ComputedRef<number>;
	/** Begin recording on `slideIndex`. */
	start: (slideIndex: number) => void;
	/** The show moved to `slideIndex`. */
	moveTo: (slideIndex: number) => void;
	togglePause: () => void;
	/** The show ended: record the current slide and open the summary. */
	finish: () => void;
	saveTimings: () => void;
	dismissSummary: () => void;
}

export function useRehearseTimings(
	options: UseRehearseTimingsOptions = {},
): UseRehearseTimingsResult {
	const controller = createPresentationRehearsalController(options);
	const state = useController(controller);
	if (getCurrentScope()) {
		onScopeDispose(controller.dispose);
	}
	return {
		rehearsing: computed(() => state.value.active),
		paused: computed(() => state.value.paused),
		showSummary: computed(() => state.value.summaryOpen),
		recordedTimings: computed(() => state.value.timings),
		slideElapsedMs: computed(() => state.value.slideElapsedMs),
		totalElapsedMs: computed(() => state.value.totalElapsedMs),
		totalRecordedMs: computed(() => state.value.recordedTotalMs),
		start: controller.start,
		moveTo: controller.moveTo,
		togglePause: controller.togglePause,
		finish: controller.finish,
		saveTimings: () => options.onSave?.({ ...controller.save() }),
		dismissSummary: controller.discard,
	};
}
