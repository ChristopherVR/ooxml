import type { CompatibilityWarningToast } from 'ooxml-ui/pptx';
import { compatibilityWarningToasts } from 'ooxml-ui/pptx';
import type { PptxCompatibilityWarning } from 'pptx-viewer-core';

/** Cap the visible toast stack; the rest collapse into a "+N" count. */
const VISIBLE_TOAST_LIMIT = 5;

/**
 * CompatToastsState: the fidelity-loss toast stack shown after a load whose
 * parse reported `PptxCompatibilityWarning`s (unmodelled markup, external
 * image references, a chart workbook that could not be written back, and so
 * on). Svelte port of Vue's `useCompatibilityToasts` composable.
 *
 * Toasts do not auto-hide (they are load diagnostics, not transient
 * notices): they persist until dismissed, and `reset()` clears them for a
 * newly loaded document, called from the same load-commit hook that resets
 * the read-only recommendation.
 */
export class CompatToastsState {
	#dismissedIds = $state<Set<string>>(new Set());
	#allDismissed = $state(false);
	/** Notices outside the load diagnostics (a failed Save); see {@link notify}. */
	#notices = $state<CompatibilityWarningToast[]>([]);
	#getWarnings: () => readonly PptxCompatibilityWarning[];

	constructor(deps: { getWarnings(): readonly PptxCompatibilityWarning[] }) {
		this.#getWarnings = deps.getWarnings;
	}

	#allToasts(): CompatibilityWarningToast[] {
		return [...this.#notices, ...compatibilityWarningToasts(this.#getWarnings())];
	}

	#activeToasts(): CompatibilityWarningToast[] {
		return this.#allDismissed
			? []
			: this.#allToasts().filter((toast) => !this.#dismissedIds.has(toast.id));
	}

	/** Every toast this load produced, dismissed ones excluded, capped at the visible limit. */
	get visibleToasts(): CompatibilityWarningToast[] {
		return this.#activeToasts().slice(0, VISIBLE_TOAST_LIMIT);
	}

	/** How many more toasts exist beyond the visible limit. */
	get overflowCount(): number {
		return Math.max(0, this.#activeToasts().length - VISIBLE_TOAST_LIMIT);
	}

	dismiss(id: string): void {
		this.#dismissedIds = new Set(this.#dismissedIds).add(id);
	}

	dismissAll(): void {
		this.#allDismissed = true;
	}

	/** Clear the dismissal state (and any notice) for a newly loaded document. */
	reset(): void {
		this.#dismissedIds = new Set();
		this.#allDismissed = false;
		this.#notices = [];
	}

	/** Show a notice (a failed Save) in the stack; shown again if it was dismissed. */
	notify(toast: CompatibilityWarningToast): void {
		// "Dismiss all" stays in force for what was already shown; only the new notice appears.
		const dismissed = this.#allDismissed
			? new Set(this.#allToasts().map((entry) => entry.id))
			: new Set(this.#dismissedIds);
		this.#notices = [toast, ...this.#notices.filter((entry) => entry.id !== toast.id)];
		dismissed.delete(toast.id);
		this.#dismissedIds = dismissed;
		this.#allDismissed = false;
	}
}
