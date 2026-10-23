import type { AiCanvasHighlight, PptxAiFocusedTarget, ToolCanvasTarget } from 'ooxml-ui/pptx/ai';
import { createPptxAiPanelController, isAiFocusPinned } from 'ooxml-ui/pptx/ai';
/**
 * The vanilla AI panel controller: the shared `createPptxAiPanelController`
 * plus a re-emit on selection changes, behind the vanilla accessor API. It owns the assistant's "scope" (focused targets +
 * a pinned override), the pick-mode state, and the on-canvas highlight sources
 * that share one overlay system:
 *
 *   - PICK MODE: the user clicks the crosshair, then clicks element(s) on the
 *     slide to hand them to the assistant. Each pick is highlighted and added
 *     to the pick set.
 *   - LIVE TOOL FOCUS: while the assistant runs its tool loop, each tool call
 *     flashes a transient "the AI is working on this" highlight on the element
 *     the tool references (see {@link flashToolTarget}).
 *
 * It is a plain observable: subscribers (the focus bar, the highlight overlay,
 * the bridge's `getFocusedTargets`) re-read state on every `subscribe` notify.
 */
import type { PptxElement } from 'pptx-viewer-core';

import type { Store, ViewerState } from '../state';

export type { AiCanvasHighlight };

/** One-shot composer prefill; `nonce` bumps on every ask/fix so it re-applies. */
export interface AiPrefill {
	text: string;
	nonce: number;
}

export interface AiFocusControllerDeps {
	store: Store<ViewerState>;
	/** Open the AI panel (used by pick / ask / fix). */
	requestOpen(): void;
}

export interface AiFocusController {
	subscribe(listener: () => void): () => void;

	/* ── Focus ──────────────────────────────────────────────────────────── */
	/** Focused targets derived live from the current canvas selection. */
	getLiveTargets(): PptxAiFocusedTarget[];
	/** Effective focus: explicit picks win over a pin, which wins over selection. */
	getEffectiveTargets(): PptxAiFocusedTarget[];
	isPinned(): boolean;
	hasPicks(): boolean;
	pinFocus(): void;
	clearPinnedFocus(): void;

	/* ── Pick mode ──────────────────────────────────────────────────────── */
	isPicking(): boolean;
	startPicking(): void;
	stopPicking(): void;
	addPick(slideIndex: number, elementId: string): void;
	clearPicks(): void;

	/* ── Live tool / canvas animation ───────────────────────────────────── */
	getHighlights(): AiCanvasHighlight[];
	/** True while the canvas should tween colour changes (AI is active). */
	isAnimating(): boolean;
	flashToolTarget(target: ToolCanvasTarget | null): void;

	/* ── Composer prefill (ask / fix) ───────────────────────────────────── */
	getPrefill(): AiPrefill;
	askAboutSelection(): void;
	fixElement(element: PptxElement | null, slideIndex: number): void;

	dispose(): void;
}

export function createAiFocusController(deps: AiFocusControllerDeps): AiFocusController {
	const { store } = deps;
	const listeners = new Set<() => void>();
	const emit = (): void => {
		for (const listener of listeners) {
			listener();
		}
	};
	const shared = createPptxAiPanelController({
		getSelection: () => {
			const state = store.get();
			return {
				activeSlideIndex: state.currentSlide,
				selectedElementIds: state.selectedElementIds,
				selectedElementId: state.selectedElementId,
			};
		},
		onOpen: () => deps.requestOpen(),
	});
	const snapshot = () => shared.getSnapshot();
	const unsubscribeShared = shared.subscribe(emit);

	// Re-emit whenever the live selection changes so the focus chips track it.
	const unsubscribeStore = store.subscribe((state, previous) => {
		if (
			state.currentSlide !== previous.currentSlide ||
			state.selectedElementId !== previous.selectedElementId ||
			state.selectedElementIds !== previous.selectedElementIds
		) {
			emit();
		}
	});

	return {
		subscribe(listener) {
			listeners.add(listener);
			return () => listeners.delete(listener);
		},
		getLiveTargets: () => shared.getLiveTargets(),
		getEffectiveTargets: () => shared.getEffectiveTargets(),
		isPinned: () => isAiFocusPinned(snapshot()),
		hasPicks: () => snapshot().pickTargets.length > 0,
		pinFocus: () => shared.pinFocus(),
		clearPinnedFocus: () => shared.clearPinnedFocus(),
		isPicking: () => snapshot().pickMode,
		startPicking: () => shared.startPicking(),
		stopPicking: () => shared.stopPicking(),
		addPick: (slideIndex, elementId) => shared.addPick(slideIndex, elementId),
		clearPicks: () => shared.clearPicks(),
		getHighlights: () => snapshot().canvasHighlights,
		isAnimating: () => snapshot().canvasAnimating,
		flashToolTarget: (target) => shared.flashToolTarget(target),
		getPrefill: () => snapshot().prefill,
		askAboutSelection: () => shared.askAboutSelection(),
		fixElement: (element, slideIndex) => shared.fixElement(element, slideIndex),
		dispose() {
			shared.dispose();
			unsubscribeShared();
			unsubscribeStore();
			listeners.clear();
		},
	};
}
