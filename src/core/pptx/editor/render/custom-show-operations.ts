/**
 * custom-show-operations: pure edits of the `p:custShowLst` model.
 *
 * Mirrors PowerPoint's Slide Show > Custom Slide Show > Custom Shows dialog:
 * New, Edit (name, add slides, remove, reorder), Copy and Remove. Every function
 * takes the current shows and returns NEW values; inputs are never mutated.
 *
 * A show lists slides by relationship id (`PptxSlide.rId`, `p:sld/@r:id`), may
 * repeat a slide (PowerPoint allows it), and is identified by an
 * `xsd:unsignedInt` id that hyperlinks and actions name with
 * `ppaction://customshow?id=N`. Slide deletion and reordering are reconciled
 * by the save pipeline (`applyCustomShows` remaps and prunes by rId); use
 * {@link pruneCustomShowSlides} to keep the in-memory model in step.
 */

import { mapCustomShowActions } from '../../index';
import type { PptxCustomShow, PptxSlide } from 'ooxml-core/pptx';

/** Result of an edit that targets one show. `show` is undefined when the id is unknown. */
export interface CustomShowEditResult {
	shows: PptxCustomShow[];
	show: PptxCustomShow | undefined;
}

/** The next free numeric `p:custShow/@id`: one above the largest numeric id in use. */
export function nextCustomShowId(shows: readonly PptxCustomShow[]): string {
	let max = -1;
	for (const show of shows) {
		const n = Number(show.id);
		if (Number.isInteger(n) && n > max) {
			max = n;
		}
	}
	return String(max + 1);
}

/** `name`, or `name (2)`, `name (3)`... when another show (not `exceptId`) already uses it. */
export function uniqueCustomShowName(
	shows: readonly PptxCustomShow[],
	name: string,
	exceptId?: string,
): string {
	const base = name.trim() || 'Custom Show';
	const taken = new Set(
		shows.filter((s) => s.id !== exceptId).map((s) => s.name.trim().toLowerCase()),
	);
	if (!taken.has(base.toLowerCase())) {
		return base;
	}
	for (let n = 2; ; n++) {
		const candidate = `${base} (${n})`;
		if (!taken.has(candidate.toLowerCase())) {
			return candidate;
		}
	}
}

function replaceShow(
	shows: readonly PptxCustomShow[],
	id: string,
	update: (show: PptxCustomShow) => PptxCustomShow,
): CustomShowEditResult {
	let edited: PptxCustomShow | undefined;
	const next = shows.map((show) => {
		if (show.id !== id) {
			return show;
		}
		edited = update(show);
		return edited;
	});
	return { shows: edited ? next : [...shows], show: edited };
}

/** Append a new show. A name already in use gets a numeric suffix. */
export function createCustomShow(
	shows: readonly PptxCustomShow[],
	name: string,
	slideRIds: readonly string[] = [],
): CustomShowEditResult {
	const show: PptxCustomShow = {
		id: nextCustomShowId(shows),
		name: uniqueCustomShowName(shows, name),
		slideRIds: [...slideRIds],
	};
	return { shows: [...shows, show], show };
}

/** Rename a show; a clash with another show's name gets a numeric suffix. */
export function renameCustomShow(
	shows: readonly PptxCustomShow[],
	id: string,
	name: string,
): CustomShowEditResult {
	return replaceShow(shows, id, (show) => ({
		...show,
		name: uniqueCustomShowName(shows, name, id),
	}));
}

/** Replace a show's whole slide list. */
export function setCustomShowSlides(
	shows: readonly PptxCustomShow[],
	id: string,
	slideRIds: readonly string[],
): CustomShowEditResult {
	return replaceShow(shows, id, (show) => ({ ...show, slideRIds: [...slideRIds] }));
}

/** Insert slides at `index` (default: the end), as the dialog's "Add" button does. */
export function addSlidesToCustomShow(
	shows: readonly PptxCustomShow[],
	id: string,
	slideRIds: readonly string[],
	index?: number,
): CustomShowEditResult {
	return replaceShow(shows, id, (show) => {
		const at = Math.max(0, Math.min(index ?? show.slideRIds.length, show.slideRIds.length));
		const list = [...show.slideRIds];
		list.splice(at, 0, ...slideRIds);
		return { ...show, slideRIds: list };
	});
}

/** Remove the entries at the given positions of a show's slide list. */
export function removeSlidesFromCustomShow(
	shows: readonly PptxCustomShow[],
	id: string,
	positions: readonly number[],
): CustomShowEditResult {
	const drop = new Set(positions);
	return replaceShow(shows, id, (show) => ({
		...show,
		slideRIds: show.slideRIds.filter((_, i) => !drop.has(i)),
	}));
}

/** Move one entry of a show's slide list (the dialog's up and down arrows). */
export function moveCustomShowSlide(
	shows: readonly PptxCustomShow[],
	id: string,
	from: number,
	to: number,
): CustomShowEditResult {
	return replaceShow(shows, id, (show) => {
		if (from < 0 || from >= show.slideRIds.length) {
			return show;
		}
		const list = [...show.slideRIds];
		const [moved] = list.splice(from, 1);
		list.splice(Math.max(0, Math.min(to, list.length)), 0, moved as string);
		return { ...show, slideRIds: list };
	});
}

/** Copy a show under a fresh id; the default name is "Copy of <name>" like PowerPoint. */
export function copyCustomShow(
	shows: readonly PptxCustomShow[],
	id: string,
	name?: string,
): CustomShowEditResult {
	const source = shows.find((s) => s.id === id);
	if (!source) {
		return { shows: [...shows], show: undefined };
	}
	// `rawXml` is not carried over: it belongs to the source's id and name.
	const show: PptxCustomShow = {
		id: nextCustomShowId(shows),
		name: uniqueCustomShowName(shows, name ?? `Copy of ${source.name}`),
		slideRIds: [...source.slideRIds],
	};
	return { shows: [...shows, show], show };
}

/** Remove a show. Use {@link clearCustomShowActions} for the actions that jump into it. */
export function deleteCustomShow(
	shows: readonly PptxCustomShow[],
	id: string,
): CustomShowEditResult {
	const show = shows.find((s) => s.id === id);
	return { shows: shows.filter((s) => s.id !== id), show };
}

/** Drop slide references whose slide is gone (after a slide delete). */
export function pruneCustomShowSlides(
	shows: readonly PptxCustomShow[],
	slides: readonly Pick<PptxSlide, 'rId'>[],
): PptxCustomShow[] {
	const live = new Set(slides.map((s) => s.rId));
	return shows.map((show) =>
		show.slideRIds.every((r) => live.has(r))
			? show
			: { ...show, slideRIds: show.slideRIds.filter((r) => live.has(r)) },
	);
}

/** The slides a show plays, in show order, skipping references that no longer resolve. */
export function customShowSlides<T extends Pick<PptxSlide, 'rId'>>(
	show: PptxCustomShow,
	slides: readonly T[],
): T[] {
	const byRId = new Map(slides.map((s) => [s.rId, s]));
	return show.slideRIds.flatMap((r) => {
		const slide = byRId.get(r);
		return slide ? [slide] : [];
	});
}

/**
 * Remove the click and hover actions and text hyperlinks that jump into custom show `id`
 * (`ppaction://customshow?id=N`), so deleting the show leaves no dangling jump.
 * Returns new slides (unchanged slides keep their identity) and how many were cleared.
 */
export function clearCustomShowActions(
	slides: readonly PptxSlide[],
	id: string,
): { slides: PptxSlide[]; cleared: number } {
	const result = mapCustomShowActions(slides, (target) => (target === id ? null : target));
	return { slides: result.slides, cleared: result.changed };
}
