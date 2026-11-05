/**
 * One back-to-front order for EVERYTHING a morph's transition overlay paints:
 * the runs of ghosts and lifted arrivals (`morph-overlay-layers`) and the
 * cross-dissolving pairs whose two halves are painted together
 * (`morph-crossfade-group`).
 *
 * Crossfade groups used to be painted above every run. That is only right
 * while nothing the overlay paints belongs above a dissolving pair, and issue
 * #141's slides 6 to 7 break it: three columns of "L"s change font (so each
 * column dissolves) while they slide under the top and bottom picture bands,
 * which the incoming slide stacks above them. Painted over every run, the
 * columns showed across both bands for the whole morph and vanished from
 * them in one frame when the overlay came down. Ranking a group at its
 * incoming shape's layer, like any other arrival, keeps the bands on top.
 *
 * A binding never sorts anything itself: it reads the z-slots from
 * {@link morphOverlayStackSlots} (one per run) or
 * {@link morphOverlayElementSlots} (one per painted element) and puts each on
 * its layer, group or element.
 *
 * @module render/morph-overlay-stack
 */
import type { PptxElement } from 'ooxml-core/pptx';

import type { MorphCrossfadeGroup } from './morph-crossfade-group';
import type { MorphOverlayLayer } from './morph-overlay-layers';
import { buildMorphIncomingAnchoredOrder } from './morph-overlay-layers';
import type { MorphPair } from './morph-types';

/** One entry of the overlay's back-to-front stack. */
export interface MorphOverlayStackEntry {
	/** `'layer'` indexes `overlayLayers`; `'crossfade'` indexes `crossfadeGroups`. */
	readonly kind: 'layer' | 'crossfade';
	readonly index: number;
}

/** The overlay's runs and the stack that interleaves them with the groups. */
export interface MorphOverlayStack {
	/** Runs of consecutive same-slide shapes, split wherever a group falls inside one. */
	readonly layers: MorphOverlayLayer[];
	/** Every run and group, back to front. */
	readonly stack: MorphOverlayStackEntry[];
}

/** The part of a morph plan the slot helpers read. */
export interface MorphOverlayStackSource {
	readonly overlayLayers?: readonly MorphOverlayLayer[];
	readonly crossfadeGroups?: readonly unknown[];
	readonly overlayStack?: readonly MorphOverlayStackEntry[];
}

type Entry =
	| { kind: 'outgoing' | 'incoming'; element: PptxElement; index: number; rank: number }
	| { kind: 'crossfade'; index: number; rank: number };

/**
 * Rank the runs' shapes and the crossfade groups in one order anchored to the
 * incoming slide (see `buildMorphIncomingAnchoredOrder`), then cut the shapes
 * into runs at every change of slide and at every group.
 *
 * A group ranks at its INCOMING half's layer, which is where the pair rests
 * once the overlay comes down. On a tie (a ghost and its own lifted
 * counterpart) the outgoing side is painted first, as `resolveMorphOverlayLayers`
 * does.
 *
 * @param ghosts - The outgoing elements the overlay paints outside groups.
 * @param lifted - The incoming elements the overlay paints outside groups.
 * @param groups - The crossfade groups, in the plan's order.
 * @param outgoing - The outgoing slide's elements, flattened, in document order.
 * @param incoming - The incoming slide's elements, flattened, in document order.
 * @param pairs - The matched pairs.
 * @param aliases - Painted copies (id -> the element they copy) that rank as
 *   the element they copy (see `morph-text-piece-plan`).
 * @returns The runs and the stack, back to front.
 */
export function resolveMorphOverlayStack(
	ghosts: readonly PptxElement[],
	lifted: readonly PptxElement[],
	groups: readonly MorphCrossfadeGroup[],
	outgoing: readonly PptxElement[],
	incoming: readonly PptxElement[],
	pairs: readonly MorphPair[],
	aliases?: ReadonlyMap<string, string>,
): MorphOverlayStack {
	const rank = buildMorphIncomingAnchoredOrder(outgoing, incoming, pairs);
	const rankOf = (id: string): number => rank.get(aliases?.get(id) ?? id) ?? 0;
	const entries: Entry[] = [
		...ghosts.map((element, index) => ({
			kind: 'outgoing' as const,
			element,
			index,
			rank: rankOf(element.id),
		})),
		...lifted.map((element, index) => ({
			kind: 'incoming' as const,
			element,
			index,
			rank: rankOf(element.id),
		})),
		...groups.map((group, index) => ({
			kind: 'crossfade' as const,
			index,
			rank: rankOf(group.incoming.id),
		})),
	];
	entries.sort(
		(a, b) =>
			a.rank - b.rank ||
			Number(a.kind !== 'outgoing') - Number(b.kind !== 'outgoing') ||
			a.index - b.index,
	);

	const layers: { side: 'outgoing' | 'incoming'; elements: PptxElement[] }[] = [];
	const stack: MorphOverlayStackEntry[] = [];
	let open: { side: 'outgoing' | 'incoming'; elements: PptxElement[] } | undefined;
	for (const entry of entries) {
		if (entry.kind === 'crossfade') {
			open = undefined;
			stack.push({ kind: 'crossfade', index: entry.index });
			continue;
		}
		if (open?.side !== entry.kind) {
			open = { side: entry.kind, elements: [] };
			stack.push({ kind: 'layer', index: layers.length });
			layers.push(open);
		}
		open.elements.push(entry.element);
	}
	return { layers, stack };
}

/** The stack a plan states, or the order before it had one: every run, then every group. */
function stackOf(plan: MorphOverlayStackSource): readonly MorphOverlayStackEntry[] {
	return (
		plan.overlayStack ?? [
			...(plan.overlayLayers ?? []).map((_, index) => ({ kind: 'layer' as const, index })),
			...(plan.crossfadeGroups ?? []).map((_, index) => ({ kind: 'crossfade' as const, index })),
		]
	);
}

/**
 * Back-to-front slots for a binding that paints each RUN as one layer:
 * `layers[i]` for `overlayLayers[i]`, `groups[j]` for `crossfadeGroups[j]`.
 * Add the binding's own base z-index to each.
 */
export function morphOverlayStackSlots(plan: MorphOverlayStackSource): {
	layers: number[];
	groups: number[];
} {
	const layers: number[] = [];
	const groups: number[] = [];
	stackOf(plan).forEach((entry, slot) => {
		(entry.kind === 'layer' ? layers : groups)[entry.index] = slot;
	});
	return { layers, groups };
}

/**
 * Back-to-front slots for a binding that gives every painted ELEMENT its own
 * z-index: each run's shapes take consecutive slots, and each group one slot
 * of its own.
 */
export function morphOverlayElementSlots(plan: MorphOverlayStackSource): {
	elements: Map<string, number>;
	groups: number[];
} {
	const elements = new Map<string, number>();
	const groups: number[] = [];
	let slot = 0;
	for (const entry of stackOf(plan)) {
		if (entry.kind === 'crossfade') {
			groups[entry.index] = slot;
			slot += 1;
			continue;
		}
		for (const element of plan.overlayLayers?.[entry.index]?.elements ?? []) {
			elements.set(element.id, slot);
			slot += 1;
		}
	}
	return { elements, groups };
}
