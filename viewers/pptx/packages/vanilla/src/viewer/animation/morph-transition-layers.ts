import type { MorphCopyPart, MorphTransitionPlan } from 'ooxml-ui/pptx';
import {
	MORPH_CROSSFADE_GROUP_CSS_TEXT,
	MORPH_CROSSFADE_HALF_BLEND_MODE,
	morphOverlayStackSlots,
} from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import { getTextBlockStyle } from '../render';
import { buildLayer } from './transition-layer';

/**
 * Strip a cloned stage down to the given element ids and drop its slide
 * background.
 *
 * Every morph layer paints a SUBSET of a slide over another layer, so it must
 * not keep that slide's own background: `getSlideBackgroundStyle` always
 * resolves to an OPAQUE fill, which would cover everything below it with a flat
 * slab for the whole transition.
 *
 * A kept shape's whole FAMILY is spared: its ancestors, and everything inside
 * it. The plan names elements at whatever level the morph matched them, so an
 * id can be a group's child (the group has to stay, or the shape this layer
 * exists to paint goes with it, and its keyframes are computed in slide space,
 * which only agrees with the DOM because children are absolutely positioned
 * inside the group's own box) or the group ITSELF (its children have to stay,
 * or the layer paints an empty box - which is what the wheel deck's centre
 * panel became the moment a whole group started dissolving as one object).
 */
export function keepOnlyElements(stage: HTMLElement, ids: readonly string[]): HTMLElement {
	stage.style.background = 'none';
	stage.style.backgroundColor = 'transparent';
	const keep = new Set(ids);
	const spared = new Set<Element>();
	for (const node of stage.querySelectorAll<HTMLElement>('[data-element-id]')) {
		const id = node.dataset.elementId;
		if (id === undefined || !keep.has(id)) {
			continue;
		}
		for (let ancestor: Element | null = node; ancestor && ancestor !== stage;) {
			spared.add(ancestor);
			ancestor = ancestor.parentElement;
		}
		for (const inside of node.querySelectorAll('[data-element-id]')) {
			spared.add(inside);
		}
	}
	for (const node of [...stage.querySelectorAll<HTMLElement>('[data-element-id]')]) {
		if (!spared.has(node)) {
			node.remove();
		}
	}
	return stage;
}

/**
 * Restack a stripped stage's kept shapes in the order the plan ranked them.
 *
 * A clone keeps every shape's z-index from its OWN slide, but a run is ordered
 * by the incoming slide (see `morph-overlay-layers`), and the two disagree
 * whenever a ghost climbs above a shape it used to sit under. Issue #141's
 * slides 9 to 10: the left column's ghost rose from under slide 9's full-slide
 * backdrop, and this binding alone kept it under that backdrop's ghost, so the
 * column vanished for the whole morph. The other bindings render the run's
 * element list, whose order is the plan's.
 */
export function stackInRunOrder(stage: HTMLElement, ids: readonly string[]): HTMLElement {
	ids.forEach((id, rank) => {
		const node = stage.querySelector<HTMLElement>(`[data-element-id="${CSS.escape(id)}"]`);
		if (node) {
			node.style.zIndex = String(rank + 1);
		}
	});
	return stage;
}

/**
 * Clone the node a morph copy stands for under the copy's own id, beside it.
 *
 * A text pair carried piece by piece (`textPieceSources`) is painted as one
 * copy of each half per piece, which the slide never rendered; the other
 * bindings render the copy's element, this one clones the stage it was given.
 * The copy can be taller than its source (its box grown to hold overflowing
 * lines), and its keyframes pivot on that box's centre, so the clone takes the
 * copy's height.
 */
export function materializeMorphCopy(
	stage: HTMLElement,
	copy: { id: string; height: number },
	sources: ReadonlyMap<string, string> | undefined,
	parts?: ReadonlyMap<string, MorphCopyPart>,
): HTMLElement {
	const id = copy.id;
	const sourceId = sources?.get(id);
	if (sourceId === undefined) {
		return stage;
	}
	const node = stage.querySelector<HTMLElement>(`[data-element-id="${CSS.escape(sourceId)}"]`);
	if (node) {
		const clone = node.cloneNode(true) as HTMLElement;
		clone.dataset.elementId = id;
		clone.style.height = `${copy.height}px`;
		keepMorphCopyPart(clone, parts?.get(id));
		if (parts?.get(id) === 'text' && 'type' in copy) {
			restyleMorphTextCopy(clone, copy as PptxElement);
		}
		node.after(clone);
	}
	return stage;
}

/**
 * A text copy stands for its source painted UNFLIPPED, a vertical flip turned
 * into a half turn (`morphUprightText`), and its frames state no flips. The
 * clone still carries the source's flip compensation on its text block (the
 * mirror that un-mirrors a flipped shape's text), so the copy's own
 * compensation (none) replaces it, or the carried text would paint mirrored.
 */
export function restyleMorphTextCopy(clone: HTMLElement, copy: PptxElement): void {
	const style = getTextBlockStyle(copy);
	for (const block of clone.querySelectorAll<HTMLElement>('.pptxv-text')) {
		block.style.transform = String(style.transform ?? '');
		block.style.transformOrigin = String(style.transformOrigin ?? '');
	}
}

/**
 * Strip a cloned shape down to the part its copy paints (`morphCopyParts`):
 * a shape whose text is carried apart from its body is painted as a copy
 * without its text and a copy without its fill and outline. The other
 * bindings render each copy from its element, which already lacks the part.
 */
export function keepMorphCopyPart(clone: HTMLElement, part: MorphCopyPart | undefined): void {
	if (part === 'body') {
		for (const text of [...clone.querySelectorAll(':scope > .pptxv-text')]) {
			text.remove();
		}
	} else if (part === 'text') {
		for (const child of [...clone.children]) {
			if (!child.classList.contains('pptxv-text')) {
				child.remove();
			}
		}
		clone.style.background = 'none';
		clone.style.border = 'none';
		clone.style.outline = 'none';
		clone.style.boxShadow = 'none';
	}
}

/** A stage clone must use the plan's destination text layout for these ghosts. */
function replaceOutgoingTextLayouts(
	stage: HTMLElement,
	incoming: HTMLElement,
	replacements: ReadonlyMap<string, string> | undefined,
): void {
	for (const [outgoingId, incomingId] of replacements ?? []) {
		const old = stage.querySelector<HTMLElement>(`[data-element-id="${CSS.escape(outgoingId)}"]`);
		const target = incoming.querySelector<HTMLElement>(
			`[data-element-id="${CSS.escape(incomingId)}"]`,
		);
		if (!old || !target) continue;
		const replacement = target.cloneNode(true) as HTMLElement;
		replacement.dataset.elementId = outgoingId;
		replacement.style.left = old.style.left;
		replacement.style.top = old.style.top;
		old.replaceWith(replacement);
	}
}

/** The z-index of the first morph overlay layer, just above the incoming layer (1). */
const MORPH_LAYER_Z_BASE = 2;

/**
 * The ghosts and lifted arrivals of a morph, one stacked layer per run of the
 * plan's `overlayLayers`, back to front.
 *
 * A ghost run is a copy of the outgoing snapshot inside a
 * `data-pptx-morph-outgoing` scope; a lifted run (arriving shapes a ghost would
 * otherwise hide, issue #146) a copy of the incoming one inside a
 * `data-pptx-morph-lifted` scope, its copy on the incoming layer held invisible
 * by the plan. One shared order, rather than every ghost under every lifted
 * shape, keeps a full-slide picture fading in from veiling the title
 * travelling above it (issue #141).
 *
 * `outgoing`/`incoming` are cloned, never modified.
 */
export function buildMorphOverlayLayers(
	doc: Document,
	outgoing: HTMLElement,
	incoming: HTMLElement,
	morphPlan: MorphTransitionPlan | undefined,
): HTMLElement[] {
	const slots = morphPlan ? morphOverlayStackSlots(morphPlan) : undefined;
	return (morphPlan?.overlayLayers ?? []).map((run, index) => {
		const source = run.side === 'outgoing' ? outgoing : incoming;
		const ids = run.elements.map((element) => element.id);
		const snapshot = source.cloneNode(true) as HTMLElement;
		if (run.side === 'outgoing') {
			replaceOutgoingTextLayouts(snapshot, incoming, morphPlan?.outgoingTextReplacements);
		}
		const layer = buildLayer(
			doc,
			stackInRunOrder(keepOnlyElements(snapshot, ids), ids),
			MORPH_LAYER_Z_BASE + (slots?.layers[index] ?? index),
			'none',
			run.side === 'outgoing' ? 'outgoing' : 'lifted',
		);
		if (run.side === 'outgoing') {
			layer.dataset.pptxMorphOutgoing = 'true';
		} else {
			layer.dataset.pptxMorphLifted = 'true';
		}
		return layer;
	});
}

/**
 * A pair the overlay paints BOTH halves of, one isolated group per pair, so
 * the two are summed rather than stacked: two source-over fades leave the
 * ink they share at 0.75 of full strength mid-transition, which bites chunks
 * out of glyphs crossing during a text dissolve, while PowerPoint's own
 * blend keeps the two coefficients summing to 1.0 (issue #161).
 *
 * `outgoing`/`incoming` must be the ORIGINAL (unstripped) stage clones: the
 * per-pair clones this function takes are cloned again internally, BEFORE
 * `playTransitionOverlay` calls {@link keepOnlyElements} on the shared
 * `outgoing` stage to build the departing morph layer.
 */
export function buildMorphCrossfadeGroups(
	doc: Document,
	outgoing: HTMLElement,
	incoming: HTMLElement,
	morphPlan: MorphTransitionPlan | undefined,
): HTMLElement[] {
	if (!morphPlan) {
		return [];
	}
	const slots = morphOverlayStackSlots(morphPlan);
	return morphPlan.crossfadeGroups.map((group, index) => {
		const wrapper = doc.createElement('div');
		wrapper.dataset.pptxMorphCrossfade = group.incoming.id;
		wrapper.style.cssText = MORPH_CROSSFADE_GROUP_CSS_TEXT;
		// `isolation` makes the wrapper a stacking context, so it needs a z-index
		// of its own: its slot in the plan's one back-to-front stack.
		wrapper.style.zIndex = String(
			MORPH_LAYER_Z_BASE + (slots.groups[index] ?? morphPlan.overlayLayers.length + index),
		);
		// The dissolve rides these WRAPPERS, not the elements: a pair dissolving in
		// place never moves, and an animation on the small element box gives it a
		// compositing layer whose raster snaps to whole device pixels, painting the
		// wording a fraction of a pixel off the live stage.
		const half = (
			stage: HTMLElement,
			element: { id: string; height: number },
			state: 'outgoing' | 'lifted',
			zIndex: number,
			animation: string | undefined,
		): HTMLElement => {
			const layer = buildLayer(
				doc,
				keepOnlyElements(
					materializeMorphCopy(
						stage,
						element,
						morphPlan.textPieceSources,
						morphPlan.morphCopyParts,
					),
					[element.id],
				),
				zIndex,
				animation ?? 'none',
				state,
			);
			if (state === 'outgoing') {
				layer.dataset.pptxMorphOutgoing = 'true';
			} else {
				layer.dataset.pptxMorphLifted = 'true';
			}
			layer.style.mixBlendMode = MORPH_CROSSFADE_HALF_BLEND_MODE;
			return layer;
		};
		wrapper.append(
			half(
				outgoing.cloneNode(true) as HTMLElement,
				group.outgoing,
				'outgoing',
				0,
				group.outgoingAnimation,
			),
			half(
				incoming.cloneNode(true) as HTMLElement,
				group.incoming,
				'lifted',
				1,
				group.incomingAnimation,
			),
		);
		return wrapper;
	});
}
