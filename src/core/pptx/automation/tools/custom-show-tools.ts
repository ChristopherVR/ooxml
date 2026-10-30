import type { PptxCustomShow, PptxData } from '../../index';
import {
	addSlidesToCustomShow,
	clearCustomShowActions,
	copyCustomShow,
	createCustomShow,
	deleteCustomShow,
	moveCustomShowSlide,
	removeSlidesFromCustomShow,
	renameCustomShow,
	setCustomShowSlides,
} from '../../editor/render/custom-show-operations';

import type { ToolContext, ToolResult } from '../types';

export interface ManageCustomShowsParams {
	action:
		| 'list'
		| 'create'
		| 'rename'
		| 'setSlides'
		| 'addSlides'
		| 'removeSlides'
		| 'moveSlide'
		| 'copy'
		| 'delete';
	showId?: string;
	name?: string;
	slideIndices?: number[];
	positions?: number[];
	from?: number;
	to?: number;
	index?: number;
}

export interface CustomShowInfo {
	id: string;
	name: string;
	/** 0-based deck indices of the show's slides, in show order (unresolved references skipped). */
	slideIndices: number[];
}

export interface ManageCustomShowsResult {
	action: string;
	shows?: CustomShowInfo[];
	show?: CustomShowInfo | null;
	/** Actions and hyperlinks that jumped into a deleted show and were removed. */
	clearedActions?: number;
}

function describeShow(show: PptxCustomShow, data: PptxData): CustomShowInfo {
	const indexByRId = new Map(data.slides.map((s, i) => [s.rId, i]));
	return {
		id: show.id,
		name: show.name,
		slideIndices: show.slideRIds.flatMap((r) => {
			const i = indexByRId.get(r);
			return i === undefined ? [] : [i];
		}),
	};
}

function rIdsFor(data: PptxData, indices: readonly number[] | undefined): string[] {
	return (indices ?? []).map((i) => {
		const slide = data.slides[i];
		if (!slide) {
			throw new Error(`Slide index ${i} out of range (0–${data.slides.length - 1}).`);
		}
		return slide.rId;
	});
}

function requireId(params: ManageCustomShowsParams): string {
	if (!params.showId) {
		throw new Error(`showId is required for ${params.action} action.`);
	}
	return params.showId;
}

/**
 * List, create, rename, edit the slides of, copy and delete custom shows
 * (`p:custShowLst`). Deleting a show also removes the click and hover actions and
 * text hyperlinks that jump into it (`ppaction://customshow?id=N`).
 */
export function manageCustomShows(
	ctx: ToolContext,
	params: ManageCustomShowsParams,
): ToolResult<ManageCustomShowsResult> {
	const data = ctx.pptxData;
	const current = data.customShows ?? [];
	const commit = (
		shows: PptxCustomShow[],
		result: ManageCustomShowsResult,
	): ToolResult<ManageCustomShowsResult> => {
		(data as { customShows?: PptxCustomShow[] }).customShows = shows;
		return { pptxData: data, dirty: true, result, saveOptions: { customShows: shows } };
	};
	const unknown = (id: string): Error => new Error(`No custom show with id ${id}.`);

	switch (params.action) {
		case 'list':
			return {
				pptxData: data,
				dirty: false,
				result: { action: 'list', shows: current.map((s) => describeShow(s, data)) },
			};

		case 'create': {
			if (!params.name?.trim()) {
				throw new Error('name is required for create action.');
			}
			const r = createCustomShow(current, params.name, rIdsFor(data, params.slideIndices));
			return commit(r.shows, { action: 'create', show: describeShow(r.show!, data) });
		}

		case 'rename': {
			if (!params.name?.trim()) {
				throw new Error('name is required for rename action.');
			}
			const id = requireId(params);
			const r = renameCustomShow(current, id, params.name);
			if (!r.show) {
				throw unknown(id);
			}
			return commit(r.shows, { action: 'rename', show: describeShow(r.show, data) });
		}

		case 'setSlides':
		case 'addSlides': {
			if (!params.slideIndices) {
				throw new Error(`slideIndices is required for ${params.action} action.`);
			}
			const id = requireId(params);
			const rIds = rIdsFor(data, params.slideIndices);
			const r =
				params.action === 'setSlides'
					? setCustomShowSlides(current, id, rIds)
					: addSlidesToCustomShow(current, id, rIds, params.index);
			if (!r.show) {
				throw unknown(id);
			}
			return commit(r.shows, { action: params.action, show: describeShow(r.show, data) });
		}

		case 'removeSlides': {
			if (!params.positions) {
				throw new Error('positions is required for removeSlides action.');
			}
			const id = requireId(params);
			const r = removeSlidesFromCustomShow(current, id, params.positions);
			if (!r.show) {
				throw unknown(id);
			}
			return commit(r.shows, { action: 'removeSlides', show: describeShow(r.show, data) });
		}

		case 'moveSlide': {
			if (params.from === undefined || params.to === undefined) {
				throw new Error('from and to are required for moveSlide action.');
			}
			const id = requireId(params);
			const r = moveCustomShowSlide(current, id, params.from, params.to);
			if (!r.show) {
				throw unknown(id);
			}
			return commit(r.shows, { action: 'moveSlide', show: describeShow(r.show, data) });
		}

		case 'copy': {
			const id = requireId(params);
			const r = copyCustomShow(current, id, params.name);
			if (!r.show) {
				throw unknown(id);
			}
			return commit(r.shows, { action: 'copy', show: describeShow(r.show, data) });
		}

		case 'delete': {
			const id = requireId(params);
			const r = deleteCustomShow(current, id);
			if (!r.show) {
				throw unknown(id);
			}
			const cleared = clearCustomShowActions(data.slides, id);
			(data as { slides: PptxData['slides'] }).slides = cleared.slides;
			const props = data.presentationProperties;
			if (props?.showSlidesCustomShowId === id) {
				props.showSlidesMode = 'all';
				delete props.showSlidesCustomShowId;
			}
			return commit(r.shows, {
				action: 'delete',
				show: describeShow(r.show, data),
				clearedActions: cleared.cleared,
			});
		}

		default:
			throw new Error(`Unknown custom show action: ${String(params.action)}`);
	}
}
