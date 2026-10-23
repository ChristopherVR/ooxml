/**
 * AiPanelStore: owns the AI panel's "scope" (focused targets + a prefilled
 * composer directive) and the two on-canvas highlight sources: the shared
 * `createPptxAiPanelController`, adapted to signals with `controllerSignal`.
 *
 * The live focus follows the canvas selection; `pinnedFocus` freezes it (set
 * from the chat or a canvas "Ask AI" affordance) so the assistant stays scoped
 * even after the user clicks away. Two highlight sources share one overlay:
 *   - PICK MODE: the user clicks a crosshair, then clicks element(s) on the
 *     slide to hand them to the assistant. Each pick is highlighted + added to
 *     {@link pickTargets}.
 *   - LIVE TOOL FOCUS: while the assistant runs its tool loop, each tool call
 *     flashes a transient "the AI is working on this" ring (see
 *     {@link flashToolTarget}).
 *
 * Provided at the viewer component level so the canvas, the panel, and the
 * bridge all read the same store.
 */
import { computed, DestroyRef, inject, Injectable } from '@angular/core';
import {
	computeFocusTargets,
	createAiChangeAnimator,
	createPptxAiPanelController,
	isAiFocusPinned,
	resolveAiEffectiveTargets,
} from 'ooxml-ui/pptx/ai';
import type {
	AiCanvasHighlight,
	AiChangeAnimationConfig,
	AiChangeAnimator,
	AiChangeBatch,
	FocusSelectionInput,
	PptxAiFocusedTarget,
	PptxAiPanelController,
	ToolCanvasTarget,
} from 'ooxml-ui/pptx/ai';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';

import { controllerSignal } from '../controller-signal';

/** Live selection accessors the store reads to derive the follow-selection focus. */
export interface AiPanelSelectionAccessors {
	activeSlideIndex(): number;
	selectedElementIds(): readonly string[];
	selectedElementId(): string | null;
	/** The primary selected element, for building the "Fix with AI" directive. */
	selectedElement(): PptxElement | null;
}

@Injectable()
export class AiPanelStore {
	/**
	 * Optional: `inject()` needs an active injection context, which plain
	 * `new AiPanelStore()` (used by this store's unit tests, deliberately
	 * bypassing TestBed for speed) does not provide. Falls back to no automatic
	 * timer cleanup when constructed outside DI, which is fine for tests.
	 */
	private readonly destroyRef: DestroyRef | null = (() => {
		try {
			return inject(DestroyRef);
		} catch {
			return null;
		}
	})();
	private accessors: AiPanelSelectionAccessors | null = null;

	/** Pinned focus override (null follows the live selection). */
	/** The shared panel controller; this store adapts its snapshot to signals. */
	private readonly controller: PptxAiPanelController = createPptxAiPanelController({
		getSelection: () => this.selection(),
		getSelectedElement: () => this.accessors?.selectedElement() ?? null,
	});
	private readonly state = controllerSignal(this.controller, this.destroyRef);

	/** Pinned focus override; null means "follow the live selection". */
	readonly pinnedFocus = computed(() => this.state().pinnedFocus);
	/** One-shot composer prefill. `nonce` bumps on every ask/fix so it re-applies. */
	readonly prefill = computed(() => this.state().prefill);
	/** True while the user is picking element(s) on the canvas for the assistant. */
	readonly pickMode = computed(() => this.state().pickMode);
	/** The elements the user has explicitly handed to the assistant (pick mode). */
	readonly pickTargets = computed(() => this.state().pickTargets);

	readonly changeAnimator: AiChangeAnimator = createAiChangeAnimator();
	/** The batch of just-applied element changes the canvas overlay animates. */
	readonly changeBatch = computed(() => this.state().changeBatch);

	constructor() {
		this.destroyRef?.onDestroy(() => {
			this.controller.dispose();
			this.changeAnimator.dispose();
		});
	}

	/** Wire the live selection accessors (called once by the viewer component). */
	bind(accessors: AiPanelSelectionAccessors): void {
		this.accessors = accessors;
	}

	private selection(): FocusSelectionInput {
		const a = this.accessors;
		return a
			? {
					activeSlideIndex: a.activeSlideIndex(),
					selectedElementIds: a.selectedElementIds(),
					selectedElementId: a.selectedElementId(),
				}
			: { activeSlideIndex: 0, selectedElementIds: [], selectedElementId: null };
	}

	/** Focused targets derived live from the current canvas selection. */
	readonly liveFocusTargets = computed<PptxAiFocusedTarget[]>(() =>
		computeFocusTargets(this.selection()),
	);

	/** True when there are explicit picks (they win over a pin / live selection). */
	readonly hasPicks = computed(() => this.pickTargets().length > 0);

	/** The targets the assistant should scope to: picks > pin > live selection. */
	readonly effectiveTargets = computed<PptxAiFocusedTarget[]>(() =>
		resolveAiEffectiveTargets(this.state(), this.liveFocusTargets()),
	);

	/** Whether a pin is currently the active focus (no picks override it). */
	readonly isPinned = computed(() => isAiFocusPinned(this.state()));

	/** Bridge accessor: the assistant's current focus scope. */
	getFocusedTargets(): PptxAiFocusedTarget[] {
		return this.effectiveTargets();
	}

	pinFocus(): void {
		this.controller.pinFocus();
	}
	clearPinnedFocus(): void {
		this.controller.clearPinnedFocus();
	}

	/** Scope to the current selection with an empty composer. */
	askAboutSelection(): void {
		this.controller.askAboutSelection();
	}

	/** Scope to the current selection with a prefilled "fix" directive. */
	fixSelection(): void {
		this.controller.fixSelection();
	}

	/* ── Pick mode ─────────────────────────────────────────────────────────── */

	/** Enter pick mode (the next canvas element clicks become picks). */
	startPicking(): void {
		this.controller.startPicking();
	}
	/** Leave pick mode without clearing the accumulated picks. */
	stopPicking(): void {
		this.controller.stopPicking();
	}
	/** Add one clicked canvas element to the pick set (and highlight it). */
	addPick(slideIndex: number, elementId: string): void {
		this.controller.addPick(slideIndex, elementId);
	}
	/** Empty the pick set and leave pick mode. */
	clearPicks(): void {
		this.controller.clearPicks();
	}

	/* ── Live tool / canvas animation ──────────────────────────────────────── */

	/**
	 * Flash a transient "the AI is working on this" highlight for a running tool,
	 * and enable colour tweening for a short settle window. Pass `null` to just
	 * enable tweening (e.g. a theme-colour edit with no single element target).
	 */
	flashToolTarget(target: ToolCanvasTarget | null): void {
		this.controller.flashToolTarget(target);
	}

	/** Element rings the canvas should draw (picks + the live tool focus). */
	readonly canvasHighlights = computed<AiCanvasHighlight[]>(() => this.state().canvasHighlights);

	/** True while the canvas should tween colour changes (AI is active). */
	readonly canvasAnimating = computed(() => this.state().canvasAnimating);

	/* ── Applied-edit change animation ─────────────────────────────────────── */

	/** Apply (or clear) the change batch the panel wants the canvas to animate. */
	showChangeBatch(batch: AiChangeBatch | null): void {
		this.controller.showChangeBatch(batch);
	}

	/**
	 * Publish an applied AI edit to the change animator (called from the bridge's
	 * write choke point with the deck slides before + after the edit). The panel's
	 * subscription reveals the slide and plays the overlay; a no-op edit or the
	 * host disabling the animation publishes nothing.
	 */
	publishAiChange(before: readonly PptxSlide[], after: readonly PptxSlide[]): void {
		this.changeAnimator.publish([...before], [...after]);
	}

	/** Apply the host's change-animation config (duration / colour / toggles). */
	configureChangeAnimation(config?: AiChangeAnimationConfig): void {
		this.changeAnimator.configure(config);
	}
}
