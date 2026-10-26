import type {
	AiCanvasHighlight,
	AiChangeBatch,
	FocusSelectionInput,
	PptxAiFocusedTarget,
	PptxAiPanelController,
	PptxAiPanelState,
	ToolCanvasTarget,
} from 'ooxml-ui/pptx/ai/host';
import {
	computeFocusTargets,
	createPptxAiPanelController,
	isAiFocusPinned,
	resolveAiEffectiveTargets,
} from 'ooxml-ui/pptx/ai/host';
/**
 * `AiPanelController` - the shared `createPptxAiPanelController` read through
 * the Svelte controller adapter, behind the runes-style accessor API. It keeps `PowerPointViewer`
 * thin by holding the focus targets, the composer prefill, and the two on-canvas
 * highlight sources that share one overlay:
 *
 *   - PICK MODE: the user hits the crosshair, then clicks element(s) on the
 *     slide to hand them to the assistant. Each pick is highlighted and added to
 *     {@link pickTargets}.
 *   - LIVE TOOL FOCUS: while the assistant runs its tool loop, each tool call
 *     flashes a transient "the AI is working on this" highlight on the element
 *     the tool references (see {@link flashToolTarget}), and enables a short
 *     colour-tween window on the canvas.
 *
 * The live focus follows the canvas selection (read through getters); `pinnedFocus`
 * freezes it (set from the chat or a canvas "Ask AI" affordance) so the assistant
 * stays scoped even after the user clicks away.
 */
import type { PptxElement } from 'pptx-viewer-core';

import { controllerState } from '../state/controller-state';
import type { ControllerState } from '../state/controller-state';

export type { AiCanvasHighlight } from 'ooxml-ui/pptx/ai/host';

/** Live viewer accessors + the panel-open hook the controller closes over. */
export interface AiPanelControllerDeps {
	getActiveSlideIndex(): number;
	getSelectedElementId(): string | null;
	getSelectedElementIds(): readonly string[];
	/** The primary selected element, for building the "Fix with AI" directive. */
	getSelectedElement(): PptxElement | undefined;
	/** Open the panel (askAboutSelection / fixSelection / startPicking need it). */
	openPanel(): void;
}

export class AiPanelController {
	readonly #deps: AiPanelControllerDeps;
	readonly #controller: PptxAiPanelController;
	readonly #state: ControllerState<PptxAiPanelState>;

	constructor(deps: AiPanelControllerDeps) {
		this.#deps = deps;
		this.#controller = createPptxAiPanelController({
			getSelection: () => this.#selection(),
			getSelectedElement: () => deps.getSelectedElement(),
			onOpen: () => deps.openPanel(),
		});
		this.#state = controllerState(this.#controller);
	}

	#selection(): FocusSelectionInput {
		return {
			activeSlideIndex: this.#deps.getActiveSlideIndex(),
			selectedElementIds: this.#deps.getSelectedElementIds(),
			selectedElementId: this.#deps.getSelectedElementId(),
		};
	}

	/** Pinned focus override (null follows the live selection). */
	get pinnedFocus(): PptxAiFocusedTarget[] | null {
		return this.#state.current.pinnedFocus;
	}

	/** One-shot composer prefill. `nonce` bumps on every ask/fix. */
	get prefill(): { text: string; nonce: number } {
		return this.#state.current.prefill;
	}

	/** True while the user is picking element(s) on the canvas for the assistant. */
	get pickMode(): boolean {
		return this.#state.current.pickMode;
	}

	/** The elements the user has explicitly handed to the assistant. */
	get pickTargets(): PptxAiFocusedTarget[] {
		return this.#state.current.pickTargets;
	}

	/** The batch of just-applied element changes the canvas should animate. */
	get changeBatch(): AiChangeBatch | null {
		return this.#state.current.changeBatch;
	}

	/** Focused targets derived live from the current canvas selection. */
	get liveFocusTargets(): PptxAiFocusedTarget[] {
		return computeFocusTargets(this.#selection());
	}

	/** True when there are explicit picks (they win over a pin / live selection). */
	get hasPicks(): boolean {
		return this.pickTargets.length > 0;
	}

	/** The targets the assistant should scope to: picks > pin > live selection. */
	get effectiveTargets(): PptxAiFocusedTarget[] {
		return resolveAiEffectiveTargets(this.#state.current, this.liveFocusTargets);
	}

	/** Whether a pin is currently the active focus (no picks override it). */
	get isPinned(): boolean {
		return isAiFocusPinned(this.#state.current);
	}

	/** Element rings the canvas should draw (picks + the live tool focus). */
	get canvasHighlights(): AiCanvasHighlight[] {
		return this.#state.current.canvasHighlights;
	}

	/** True while the canvas should tween colour changes (AI is active). */
	get canvasAnimating(): boolean {
		return this.#state.current.canvasAnimating;
	}

	pinFocus(): void {
		this.#controller.pinFocus();
	}

	clearPinnedFocus(): void {
		this.#controller.clearPinnedFocus();
	}

	/** Open the panel scoped to the current selection, empty composer (focused). */
	askAboutSelection(): void {
		this.#controller.askAboutSelection();
	}

	/** Open the panel scoped to the current selection, prefilled fix directive. */
	fixSelection(): void {
		this.#controller.fixSelection();
	}

	/** Enter pick mode (the next canvas element clicks become picks). */
	startPicking(): void {
		this.#controller.startPicking();
	}

	/** Leave pick mode without clearing the accumulated picks. */
	stopPicking(): void {
		this.#controller.stopPicking();
	}

	/** Add one clicked canvas element to the pick set (and highlight it). */
	addPick(slideIndex: number, elementId: string): void {
		this.#controller.addPick(slideIndex, elementId);
	}

	/** Empty the pick set and leave pick mode. */
	clearPicks(): void {
		this.#controller.clearPicks();
	}

	/**
	 * Flash a transient "the AI is working on this" highlight for a running tool,
	 * and enable colour tweening for a short settle window. Pass `null` to just
	 * enable tweening (e.g. a theme-colour edit with no single element target).
	 */
	flashToolTarget(target: ToolCanvasTarget | null): void {
		this.#controller.flashToolTarget(target);
	}

	/**
	 * Push (or clear) the change batch the AI apply path published, so the canvas
	 * overlay animates the just-applied edit. Pass `null` to clear once settled.
	 */
	showChangeBatch(batch: AiChangeBatch | null): void {
		this.#controller.showChangeBatch(batch);
	}

	/** Cancel the pending flash timer (call on teardown). */
	dispose(): void {
		this.#controller.dispose();
	}
}
