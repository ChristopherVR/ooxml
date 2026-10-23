import type {
	AiCanvasHighlight,
	AiChangeBatch,
	FocusSelectionInput,
	PptxAiFocusedTarget,
	ToolCanvasTarget,
} from 'ooxml-ui/pptx/ai';
import { computeFocusTargets, createPptxAiPanelController } from 'ooxml-ui/pptx/ai';
/**
 * useAiPanelController: owns the AI panel's "scope" (the focused targets + a
 * prefilled composer directive) plus the two on-canvas highlight sources that
 * share one overlay system: the shared `createPptxAiPanelController`, mirrored
 * into Vue with the generic `useController` adapter. The panel open state itself stays in
 * `PowerPointViewer` (an `aiPanelOpen` ref), so this composable only owns focus,
 * picks and the live-tool flash.
 *
 *   - PICK MODE: the user clicks the crosshair button, then clicks element(s) on
 *     the slide to hand them to the assistant. Each pick is highlighted and
 *     added to {@link pickTargets}.
 *   - LIVE TOOL FOCUS: while the assistant runs its tool loop, each tool call
 *     flashes a transient "the AI is working on this" highlight on the element
 *     the tool references (see {@link flashToolTarget}).
 */
import type { PptxElement } from 'pptx-viewer-core';
import { computed, onScopeDispose } from 'vue';
import type { ComputedRef, Ref } from 'vue';

import { useController } from '../useController';

export type { AiCanvasHighlight } from 'ooxml-ui/pptx/ai';

/** Live reactive inputs the controller derives its focus from. */
export interface UseAiPanelControllerInput {
	activeSlideIndex: Ref<number>;
	selectedElementIds: Ref<string[]>;
	/** The primary selected element, for building the "Fix with AI" directive. */
	selectedElement: () => PptxElement | null;
}

export interface AiPanelController {
	/** Focused targets derived live from the current canvas selection. */
	liveFocusTargets: ComputedRef<PptxAiFocusedTarget[]>;
	/** Pinned focus override (null follows the live selection). */
	pinnedFocus: Readonly<Ref<PptxAiFocusedTarget[] | null>>;
	/** Pin the current live targets as the chat's focus. */
	pinFocus(): void;
	/** Drop the pin and follow the live selection again. */
	clearPinnedFocus(): void;
	/** One-shot composer prefill; `nonce` bumps on every ask/fix. */
	prefill: Readonly<Ref<{ text: string; nonce: number }>>;
	/** Scope the panel to the current selection, empty composer (bumps prefill). */
	askAboutSelection(): void;
	/** Scope the panel to the current selection, prefilled fix directive. */
	fixSelection(): void;

	/* Pick mode */
	pickMode: Readonly<Ref<boolean>>;
	startPicking(): void;
	stopPicking(): void;
	pickTargets: Readonly<Ref<PptxAiFocusedTarget[]>>;
	addPick(slideIndex: number, elementId: string): void;
	clearPicks(): void;

	/* Live tool / canvas animation */
	canvasHighlights: ComputedRef<AiCanvasHighlight[]>;
	canvasAnimating: ComputedRef<boolean>;
	flashToolTarget(target: ToolCanvasTarget | null): void;

	/* Applied-edit change animation */
	/** The batch of just-applied element changes the canvas should animate. */
	changeBatch: Readonly<Ref<AiChangeBatch | null>>;
	/** Push (or clear) the change batch the AI apply path published. */
	showChangeBatch(batch: AiChangeBatch | null): void;
}

/** How long a live-tool highlight / colour-tween window stays up after a call. */
export function useAiPanelController(input: UseAiPanelControllerInput): AiPanelController {
	const { activeSlideIndex, selectedElementIds, selectedElement } = input;
	const selection = (): FocusSelectionInput => ({
		activeSlideIndex: activeSlideIndex.value,
		selectedElementIds: selectedElementIds.value,
		selectedElementId: selectedElementIds.value[0] ?? null,
	});
	// The shared controller owns the state; `useController` mirrors its snapshot.
	const controller = createPptxAiPanelController({
		getSelection: selection,
		getSelectedElement: () => selectedElement(),
	});
	const state = useController(controller);
	onScopeDispose(() => controller.dispose());

	return {
		liveFocusTargets: computed(() => computeFocusTargets(selection())),
		pinnedFocus: computed(() => state.value.pinnedFocus),
		pinFocus: () => controller.pinFocus(),
		clearPinnedFocus: () => controller.clearPinnedFocus(),
		prefill: computed(() => state.value.prefill),
		askAboutSelection: () => controller.askAboutSelection(),
		fixSelection: () => controller.fixSelection(),
		pickMode: computed(() => state.value.pickMode),
		startPicking: () => controller.startPicking(),
		stopPicking: () => controller.stopPicking(),
		pickTargets: computed(() => state.value.pickTargets),
		addPick: (slideIndex, elementId) => controller.addPick(slideIndex, elementId),
		clearPicks: () => controller.clearPicks(),
		canvasHighlights: computed(() => state.value.canvasHighlights),
		canvasAnimating: computed(() => state.value.canvasAnimating),
		flashToolTarget: (target) => controller.flashToolTarget(target),
		changeBatch: computed(() => state.value.changeBatch),
		showChangeBatch: (batch) => controller.showChangeBatch(batch),
	};
}
