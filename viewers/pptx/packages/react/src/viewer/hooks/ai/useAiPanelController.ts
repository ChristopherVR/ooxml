import type { AiChangeBatch, PptxAiFocusedTarget, ToolCanvasTarget } from 'ooxml-ui/pptx/ai/host';
import { computeFocusTargets, createPptxAiPanelController } from 'ooxml-ui/pptx/ai/host';
/**
 * useAiPanelController: owns the AI panel's open state and its "scope" (the
 * focused targets + a prefilled composer directive) so PowerPointViewer can
 * stay thin. The live focus follows the canvas selection; `pinnedFocus` freezes
 * it (set from the chat or from a canvas "Ask AI" affordance) so the assistant
 * stays scoped even after the user clicks away.
 *
 * It also owns two on-canvas highlight sources that share one overlay system:
 *   - PICK MODE: the user clicks a target/crosshair button, then clicks
 *     element(s) on the slide to hand them to the assistant. Each pick is
 *     highlighted and added to {@link pickTargets}.
 *   - LIVE TOOL FOCUS: while the assistant runs its tool loop, each tool call
 *     flashes a transient "the AI is working on this" highlight on the element
 *     the tool references (see {@link flashToolTarget}).
 */
import type { PptxElement } from 'pptx-viewer-core';
import { useEffect, useMemo, useRef, useState } from 'react';

import { useController } from '../useController';

export interface UseAiPanelControllerInput {
	activeSlideIndex: number;
	selectedElementId: string | null;
	selectedElementIds: string[];
	/** The primary selected element, for building the "Fix with AI" directive. */
	selectedElement: PptxElement | null;
}

/** One element ring to draw on the canvas: an explicit pick or a live-tool focus. */
export interface AiCanvasHighlight {
	slideIndex: number;
	elementId: string;
	/** `pick` = persistent user pick; `active` = transient AI-is-working ring. */
	variant: 'pick' | 'active';
}

export interface AiPanelController {
	isOpen: boolean;
	open(): void;
	close(): void;
	toggle(): void;
	/** Focused targets derived live from the current canvas selection. */
	liveFocusTargets: PptxAiFocusedTarget[];
	/** Pinned focus override (null follows the live selection). */
	pinnedFocus: PptxAiFocusedTarget[] | null;
	/** Pin the current live targets as the chat's focus. */
	pinFocus(): void;
	/** Drop the pin and follow the live selection again. */
	clearPinnedFocus(): void;
	/**
	 * One-shot composer prefill. `nonce` bumps on every ask/fix so the composer
	 * focuses (and applies `text`) even when `text` is empty and unchanged.
	 */
	prefill: { text: string; nonce: number };
	/** Open the panel scoped to the current selection, empty composer (focused). */
	askAboutSelection(): void;
	/** Open the panel scoped to the current selection, prefilled fix directive. */
	fixSelection(): void;

	/* ── Pick mode ─────────────────────────────────────────────────────────── */
	/** True while the user is picking element(s) on the canvas for the assistant. */
	pickMode: boolean;
	/** Enter pick mode (the next canvas element clicks become picks). */
	startPicking(): void;
	/** Leave pick mode without clearing the accumulated picks. */
	stopPicking(): void;
	/** The elements the user has explicitly handed to the assistant. */
	pickTargets: PptxAiFocusedTarget[];
	/** Add one clicked canvas element to the pick set (and highlight it). */
	addPick(slideIndex: number, elementId: string): void;
	/** Empty the pick set and leave pick mode. */
	clearPicks(): void;

	/* ── Live tool / canvas animation ──────────────────────────────────────── */
	/** Element rings the canvas should draw (picks + the live tool focus). */
	canvasHighlights: AiCanvasHighlight[];
	/** True while the canvas should tween colour changes (AI is active). */
	canvasAnimating: boolean;
	/**
	 * Flash a transient "the AI is working on this" highlight for a running tool,
	 * and enable colour tweening for a short settle window. Pass `null` to just
	 * enable tweening (e.g. a theme-colour edit with no single element target).
	 */
	flashToolTarget(target: ToolCanvasTarget | null): void;

	/* ── Applied-edit change animation ──────────────────────────────────────── */
	/** The batch of just-applied element changes the canvas should animate. */
	changeBatch: AiChangeBatch | null;
	/** Push (or clear) the change batch the AI apply path published. */
	showChangeBatch(batch: AiChangeBatch | null): void;
}

export function useAiPanelController(input: UseAiPanelControllerInput): AiPanelController {
	const { activeSlideIndex, selectedElementId, selectedElementIds } = input;
	const inputRef = useRef(input);
	inputRef.current = input;
	// The shared controller owns the state; `useController` re-renders on change.
	const [controller] = useState(() =>
		createPptxAiPanelController({
			getSelection: () => ({
				activeSlideIndex: inputRef.current.activeSlideIndex,
				selectedElementIds: inputRef.current.selectedElementIds,
				selectedElementId: inputRef.current.selectedElementId,
			}),
			getSelectedElement: () => inputRef.current.selectedElement,
		}),
	);
	const state = useController(controller);
	useEffect(() => () => controller.dispose(), [controller]);

	const liveFocusTargets = useMemo(
		() => computeFocusTargets({ activeSlideIndex, selectedElementIds, selectedElementId }),
		[activeSlideIndex, selectedElementIds, selectedElementId],
	);

	// Stable action identities: they only call into the controller.
	const actions = useMemo(
		() => ({
			open: () => controller.open(),
			close: () => controller.close(),
			toggle: () => controller.toggle(),
			pinFocus: () => controller.pinFocus(),
			clearPinnedFocus: () => controller.clearPinnedFocus(),
			askAboutSelection: () => controller.askAboutSelection(),
			fixSelection: () => controller.fixSelection(),
			startPicking: () => controller.startPicking(),
			stopPicking: () => controller.stopPicking(),
			addPick: (slideIndex: number, elementId: string) => controller.addPick(slideIndex, elementId),
			clearPicks: () => controller.clearPicks(),
			flashToolTarget: (target: ToolCanvasTarget | null) => controller.flashToolTarget(target),
			showChangeBatch: (batch: AiChangeBatch | null) => controller.showChangeBatch(batch),
		}),
		[controller],
	);

	return useMemo<AiPanelController>(
		() => ({
			...actions,
			isOpen: state.isOpen,
			liveFocusTargets,
			pinnedFocus: state.pinnedFocus,
			prefill: state.prefill,
			pickMode: state.pickMode,
			pickTargets: state.pickTargets,
			canvasHighlights: state.canvasHighlights,
			canvasAnimating: state.canvasAnimating,
			changeBatch: state.changeBatch,
		}),
		[actions, liveFocusTargets, state],
	);
}
