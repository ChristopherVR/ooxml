import type { CompatibilityWarningToast } from 'ooxml-ui/pptx';
import { compatibilityWarningToasts } from 'ooxml-ui/pptx';
import type { PptxCompatibilityWarning } from 'pptx-viewer-core';
import { computed, ref } from 'vue';
import type { ComputedRef, Ref } from 'vue';

/** Cap the visible toast stack; the rest collapse into a "+N" count. */
const VISIBLE_TOAST_LIMIT = 5;

/**
 * useCompatibilityToasts: the fidelity-loss toast stack shown after a load
 * whose parse reported `PptxCompatibilityWarning`s (unmodelled markup,
 * external image references, a chart workbook that could not be written
 * back, and so on). Every warning already flows into
 * `useLoadContent().compatibilityWarnings`; this composable only maps that
 * list through the shared decision function and tracks per-toast /
 * dismiss-all state.
 *
 * Toasts do not auto-hide (they are load diagnostics, not transient
 * notices): they persist until dismissed, and `reset()` clears them for a
 * newly loaded document, called from the same place `useLoadContent`'s
 * consumer resets Protected View / the read-only recommendation.
 */
export interface UseCompatibilityToastsInput {
	warnings: Ref<PptxCompatibilityWarning[]>;
}

export interface UseCompatibilityToastsResult {
	/** Every toast this load produced, most-recently-relevant first, dismissed ones excluded. */
	visibleToasts: ComputedRef<CompatibilityWarningToast[]>;
	/** How many more toasts exist beyond {@link VISIBLE_TOAST_LIMIT}. */
	overflowCount: ComputedRef<number>;
	dismiss: (id: string) => void;
	dismissAll: () => void;
	reset: () => void;
	/**
	 * Show a notice outside the load diagnostics (a failed Save); shown again if it was dismissed.
	 * Cleared by {@link reset} with the rest.
	 */
	notify: (toast: CompatibilityWarningToast) => void;
}

export function useCompatibilityToasts(
	input: UseCompatibilityToastsInput,
): UseCompatibilityToastsResult {
	const dismissedIds = ref<Set<string>>(new Set());
	const allDismissed = ref(false);

	const notices = ref<CompatibilityWarningToast[]>([]);
	const allToasts = computed(() => [
		...notices.value,
		...compatibilityWarningToasts(input.warnings.value),
	]);

	const activeToasts = computed(() =>
		allDismissed.value ? [] : allToasts.value.filter((toast) => !dismissedIds.value.has(toast.id)),
	);

	const visibleToasts = computed(() => activeToasts.value.slice(0, VISIBLE_TOAST_LIMIT));
	const overflowCount = computed(() =>
		Math.max(0, activeToasts.value.length - VISIBLE_TOAST_LIMIT),
	);

	function dismiss(id: string): void {
		dismissedIds.value = new Set(dismissedIds.value).add(id);
	}

	function dismissAll(): void {
		allDismissed.value = true;
	}

	function reset(): void {
		dismissedIds.value = new Set();
		allDismissed.value = false;
		notices.value = [];
	}

	function notify(toast: CompatibilityWarningToast): void {
		// "Dismiss all" stays in force for what was already shown; only the new notice appears.
		const dismissed = allDismissed.value
			? new Set(allToasts.value.map((entry) => entry.id))
			: new Set(dismissedIds.value);
		notices.value = [toast, ...notices.value.filter((entry) => entry.id !== toast.id)];
		dismissed.delete(toast.id);
		dismissedIds.value = dismissed;
		allDismissed.value = false;
	}

	return { visibleToasts, overflowCount, dismiss, dismissAll, reset, notify };
}
