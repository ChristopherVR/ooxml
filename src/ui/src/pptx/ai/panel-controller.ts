/**
 * The AI panel controller: the assistant's scope (focused targets and a pinned
 * override), pick mode, the composer prefill and the two on-canvas highlight
 * sources that share one overlay.
 *
 *   - PICK MODE: the user clicks the crosshair, then clicks element(s) on the
 *     slide to hand them to the assistant. Each pick is highlighted.
 *   - LIVE TOOL FOCUS: while the assistant runs its tool loop, each tool call
 *     flashes a transient "the AI is working on this" ring on the element the
 *     tool references (see `flashToolTarget`) and opens a short colour-tween
 *     window on the canvas.
 *
 * All five bindings carried a port of this (React and Vue hooks, an Angular
 * store, a Svelte runes class, a vanilla observable) with the same behaviour;
 * it follows the controller convention (`getSnapshot` / `subscribe` / typed
 * actions) in `render/controller`, so each binding only adapts the snapshot to
 * its reactivity. The live selection stays binding state: the controller reads
 * it through `deps` when it pins, and {@link resolveAiEffectiveTargets} derives
 * the effective focus from a snapshot plus the binding's live targets.
 *
 * @module ai/panel-controller
 */

import type { PptxElement } from 'ooxml-core/pptx';

import type { PptxController } from '../render/controller/controller';
import { createControllerStore } from '../render/controller/controller';
import type { PptxAiFocusedTarget } from './bridge';
import type { AiChangeBatch } from './change-animator';
import type { AiCanvasHighlight, FocusSelectionInput } from './focus-targets';
import { computeFocusTargets } from './focus-targets';
import type { ToolCanvasTarget } from './tool-target';

/** How long a live-tool highlight / colour-tween window stays up after a call. */
export const AI_TOOL_FLASH_MS = 2600;

/** One-shot composer prefill; `nonce` bumps on every ask / fix so it re-applies. */
export interface PptxAiPrefill {
	text: string;
	nonce: number;
}

/** The panel controller's state. */
export interface PptxAiPanelState {
	/** The panel is open (bindings that own the open state elsewhere ignore it). */
	readonly isOpen: boolean;
	/** Pinned focus override; null follows the live selection. */
	readonly pinnedFocus: PptxAiFocusedTarget[] | null;
	readonly prefill: PptxAiPrefill;
	/** The user is picking element(s) on the canvas for the assistant. */
	readonly pickMode: boolean;
	/** Elements the user handed to the assistant; they win over a pin. */
	readonly pickTargets: PptxAiFocusedTarget[];
	/** Element rings the canvas draws (picks, then the live tool focus). */
	readonly canvasHighlights: AiCanvasHighlight[];
	/** True while the canvas should tween colour changes (AI is active). */
	readonly canvasAnimating: boolean;
	/** The just-applied change batch the canvas overlay animates. */
	readonly changeBatch: AiChangeBatch | null;
}

/** Live viewer reads the controller needs when it pins the current focus. */
export interface PptxAiPanelControllerDeps {
	/** The live selection (active slide, selected ids). */
	getSelection(): FocusSelectionInput;
	/** The primary selected element, for the "Fix with AI" directive. */
	getSelectedElement?(): PptxElement | null | undefined;
	/** Called whenever an action opens the panel (pick, ask, fix, `open`). */
	onOpen?(): void;
}

/** The AI panel controller (see the module note). */
export interface PptxAiPanelController extends PptxController<PptxAiPanelState> {
	open(): void;
	close(): void;
	toggle(): void;
	/** Focused targets derived from the live selection right now. */
	getLiveTargets(): PptxAiFocusedTarget[];
	/** The assistant's focus right now: picks, else a pin, else the live selection. */
	getEffectiveTargets(): PptxAiFocusedTarget[];
	pinFocus(): void;
	clearPinnedFocus(): void;
	/** Open scoped to the current selection with an empty composer. */
	askAboutSelection(): void;
	/** Open scoped to the current selection with the fix directive for the selected element. */
	fixSelection(): void;
	/** Open scoped to the current selection with the fix directive for `element`. */
	fixElement(element: PptxElement | null | undefined, slideIndex: number): void;
	startPicking(): void;
	/** Leave pick mode, keeping the picks. */
	stopPicking(): void;
	addPick(slideIndex: number, elementId: string): void;
	/** Empty the picks and leave pick mode. */
	clearPicks(): void;
	/**
	 * Flash the "AI is working on this" ring for a running tool and open a short
	 * colour-tween window. `null` only opens the window (a theme edit, say).
	 */
	flashToolTarget(target: ToolCanvasTarget | null): void;
	/** Publish (or clear with `null`) the change batch the canvas animates. */
	showChangeBatch(batch: AiChangeBatch | null): void;
	/** Cancel the pending flash timer (call on teardown). */
	dispose(): void;
}

/** The "Fix with AI" directive for one element (prefilled, never auto-sent). */
export function aiFixDirective(element: PptxElement, slideIndex: number): string {
	return `Review and fix any issues with this ${element.type} (id=${element.id}) on slide ${slideIndex + 1}.`;
}

/** The assistant's effective focus: picks win over a pin, which wins over the live selection. */
export function resolveAiEffectiveTargets(
	state: Pick<PptxAiPanelState, 'pinnedFocus' | 'pickTargets'>,
	liveTargets: PptxAiFocusedTarget[],
): PptxAiFocusedTarget[] {
	if (state.pickTargets.length > 0) {
		return state.pickTargets;
	}
	if (state.pinnedFocus && state.pinnedFocus.length > 0) {
		return state.pinnedFocus;
	}
	return liveTargets;
}

/** A pin is the active focus (no picks override it). */
export function isAiFocusPinned(
	state: Pick<PptxAiPanelState, 'pinnedFocus' | 'pickTargets'>,
): boolean {
	return state.pickTargets.length === 0 && state.pinnedFocus !== null;
}

interface ToolFocus {
	slideIndex: number;
	elementIds: string[];
}

function buildHighlights(
	pickTargets: PptxAiFocusedTarget[],
	toolFocus: ToolFocus | null,
): AiCanvasHighlight[] {
	const out: AiCanvasHighlight[] = [];
	for (const target of pickTargets) {
		if (target.kind === 'element') {
			out.push({ slideIndex: target.slideIndex, elementId: target.elementId, variant: 'pick' });
		}
	}
	for (const elementId of toolFocus?.elementIds ?? []) {
		out.push({ slideIndex: toolFocus?.slideIndex ?? 0, elementId, variant: 'active' });
	}
	return out;
}

const INITIAL_STATE: PptxAiPanelState = {
	isOpen: false,
	pinnedFocus: null,
	prefill: { text: '', nonce: 0 },
	pickMode: false,
	pickTargets: [],
	canvasHighlights: [],
	canvasAnimating: false,
	changeBatch: null,
};

/** Build the AI panel controller over the binding's live selection. */
export function createPptxAiPanelController(
	deps: PptxAiPanelControllerDeps,
): PptxAiPanelController {
	const store = createControllerStore<PptxAiPanelState>(INITIAL_STATE);
	let toolFocus: ToolFocus | null = null;
	let flashing = false;
	let flashTimer: ReturnType<typeof setTimeout> | null = null;

	const update = (patch: Partial<PptxAiPanelState>): void => {
		store.update((current) => {
			const next = { ...current, ...patch };
			const canvasHighlights = buildHighlights(next.pickTargets, toolFocus);
			const sameHighlights =
				canvasHighlights.length === current.canvasHighlights.length &&
				canvasHighlights.every((h, i) => {
					const prev = current.canvasHighlights[i];
					return (
						prev?.elementId === h.elementId &&
						prev.slideIndex === h.slideIndex &&
						prev.variant === h.variant
					);
				});
			return {
				...next,
				canvasHighlights: sameHighlights ? current.canvasHighlights : canvasHighlights,
				canvasAnimating: flashing || canvasHighlights.length > 0,
			};
		});
	};

	const liveTargets = (): PptxAiFocusedTarget[] => computeFocusTargets(deps.getSelection());

	const openPanel = (patch: Partial<PptxAiPanelState>): void => {
		update({ ...patch, isOpen: true });
		deps.onOpen?.();
	};

	const fixElement = (element: PptxElement | null | undefined, slideIndex: number): void => {
		const { prefill } = store.getSnapshot();
		openPanel({
			pinnedFocus: liveTargets(),
			prefill: {
				text: element ? aiFixDirective(element, slideIndex) : '',
				nonce: prefill.nonce + 1,
			},
		});
	};

	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		open: () => openPanel({}),
		close: () => update({ isOpen: false }),
		toggle() {
			if (store.getSnapshot().isOpen) {
				update({ isOpen: false });
			} else {
				openPanel({});
			}
		},
		getLiveTargets: liveTargets,
		getEffectiveTargets: () => resolveAiEffectiveTargets(store.getSnapshot(), liveTargets()),
		pinFocus: () => update({ pinnedFocus: liveTargets() }),
		clearPinnedFocus: () => update({ pinnedFocus: null }),
		askAboutSelection() {
			const { prefill } = store.getSnapshot();
			openPanel({ pinnedFocus: liveTargets(), prefill: { text: '', nonce: prefill.nonce + 1 } });
		},
		fixSelection() {
			fixElement(deps.getSelectedElement?.(), deps.getSelection().activeSlideIndex);
		},
		fixElement,
		startPicking: () => openPanel({ pickMode: true }),
		stopPicking: () => update({ pickMode: false }),
		addPick(slideIndex, elementId) {
			const { pickTargets } = store.getSnapshot();
			if (pickTargets.some((t) => t.kind === 'element' && t.elementId === elementId)) {
				return;
			}
			update({ pickTargets: [...pickTargets, { kind: 'element', slideIndex, elementId }] });
		},
		clearPicks: () => update({ pickTargets: [], pickMode: false }),
		flashToolTarget(target) {
			toolFocus =
				target && target.slideIndex !== undefined && target.elementIds.length > 0
					? { slideIndex: target.slideIndex, elementIds: target.elementIds }
					: null;
			flashing = true;
			if (flashTimer) {
				clearTimeout(flashTimer);
			}
			flashTimer = setTimeout(() => {
				toolFocus = null;
				flashing = false;
				flashTimer = null;
				update({});
			}, AI_TOOL_FLASH_MS);
			update({});
		},
		showChangeBatch: (batch) => update({ changeBatch: batch }),
		dispose() {
			if (flashTimer) {
				clearTimeout(flashTimer);
				flashTimer = null;
			}
		},
	};
}
