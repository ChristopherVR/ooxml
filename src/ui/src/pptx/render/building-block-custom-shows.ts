import type { PptxCustomShow, PptxSlide } from 'ooxml-core/pptx';
import {
	customShowSlides,
	setCustomShowSlides,
} from 'ooxml-core/pptx/editor/render/custom-show-operations';

import type { BuildingBlockFailureCode } from './building-block-actions';
import { createCustomShowsController } from './controller/custom-shows-controller';

/**
 * Framework-free edits behind Slide Show > Custom Slide Show > Custom Shows for a host that
 * renders its own list (`useViewerBuildingBlocks().actions.customShows`). Each edit runs on the
 * same headless controller the bundled dialog uses, so naming (a clash gets a numeric suffix),
 * ids and the clean-up of jumps into a removed show are identical; the host applies it at once
 * instead of through an OK button.
 */

/** One edit of the deck's custom shows. */
export type CustomShowEdit =
	| { kind: 'create'; name: string; slideRIds?: readonly string[] }
	| { kind: 'rename'; id: string; name: string }
	| { kind: 'setSlides'; id: string; slideRIds: readonly string[] }
	| { kind: 'copy'; id: string; name?: string }
	| { kind: 'remove'; id: string };

/** What storing an edit needs, mirroring the Custom Shows dialog's OK. */
export interface CustomShowEditOutcome {
	shows: PptxCustomShow[];
	/** The slides with jumps into a removed show cleared, undefined when none changed. */
	slides: PptxSlide[] | undefined;
	/** The show the edit created or changed (undefined for a removal). */
	showId: string | undefined;
}

export type CustomShowEditResult =
	| ({ ok: true } & CustomShowEditOutcome)
	| { ok: false; code: BuildingBlockFailureCode };

/** Apply `edit` to `shows`; refused with a code when it names an unknown show or slide. */
export function applyCustomShowEdit(
	shows: readonly PptxCustomShow[],
	slides: readonly PptxSlide[],
	edit: CustomShowEdit,
): CustomShowEditResult {
	const known = new Set(slides.map((slide) => slide.rId));
	const fail = (code: BuildingBlockFailureCode) => ({ ok: false as const, code });
	if ('id' in edit && !shows.some((show) => show.id === edit.id)) {
		return fail('unknown-show');
	}
	if ((edit.kind === 'create' || edit.kind === 'rename') && edit.name.trim() === '') {
		return fail('invalid-name');
	}
	const rIds = edit.kind === 'create' || edit.kind === 'setSlides' ? edit.slideRIds : undefined;
	if (rIds?.some((rId) => !known.has(rId))) {
		return fail('unknown-slide');
	}
	const controller = createCustomShowsController(shows, 'id' in edit ? edit.id : null);
	let showId: string | undefined = 'id' in edit ? edit.id : undefined;
	let next: readonly PptxCustomShow[];
	if (edit.kind === 'setSlides') {
		next = setCustomShowSlides(shows, edit.id, edit.slideRIds).shows;
	} else {
		if (edit.kind === 'create') {
			controller.create(edit.name);
		} else if (edit.kind === 'rename') {
			controller.rename(edit.name);
		} else if (edit.kind === 'copy') {
			controller.copySelected(edit.name);
		} else {
			controller.removeSelected();
		}
		const state = controller.getSnapshot();
		next = state.shows;
		showId = edit.kind === 'remove' ? undefined : (state.selectedId ?? undefined);
	}
	if (edit.kind === 'create' && rIds && rIds.length > 0 && showId !== undefined) {
		next = setCustomShowSlides(next, showId, rIds).shows;
	}
	const result = controller.result(slides);
	return {
		ok: true,
		shows: next.map((show) => ({ ...show, slideRIds: [...show.slideRIds] })),
		slides: result.slides,
		showId,
	};
}

/** The slides, in show order, that custom show `id` would play; refused when it is unknown or empty. */
export function resolveCustomShowStart(
	shows: readonly PptxCustomShow[],
	slides: readonly PptxSlide[],
	id: string,
): { ok: true; firstSlideIndex: number } | { ok: false; code: BuildingBlockFailureCode } {
	const show = shows.find((candidate) => candidate.id === id);
	if (!show) {
		return { ok: false, code: 'unknown-show' };
	}
	const first = customShowSlides(show, slides)[0];
	return first
		? { ok: true, firstSlideIndex: slides.indexOf(first) }
		: { ok: false, code: 'no-slide' };
}

/** A custom show as a host's list needs it. */
export interface CustomShowSummary {
	id: string;
	name: string;
	/** The slide list as relationship ids, in show order (a slide may repeat). */
	slideRIds: string[];
	/** How many of those still resolve to a slide of the deck. */
	slideCount: number;
}

/** A slide a custom show may list. */
export interface CustomShowSlideChoice {
	rId: string;
	/** 1-based number in the deck. */
	number: number;
}

export function describeCustomShows(
	shows: readonly PptxCustomShow[],
	slides: readonly PptxSlide[],
): { shows: CustomShowSummary[]; slides: CustomShowSlideChoice[] } {
	return {
		shows: shows.map((show) => ({
			id: show.id,
			name: show.name,
			slideRIds: [...show.slideRIds],
			slideCount: customShowSlides(show, slides).length,
		})),
		slides: slides.map((slide, index) => ({ rId: slide.rId, number: index + 1 })),
	};
}
