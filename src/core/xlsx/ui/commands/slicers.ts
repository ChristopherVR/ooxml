// Slicer selection helpers shared by the Insert > Slicer command, the Slicer tab and the grid. A
// slicer is "active" while its frame is the selected drawing (clicked in the grid).
import {
	pivotTableAt,
	type PivotTable,
	type Slicer,
	type SlicerCache,
	type SlicerFrameObject,
	type Timeline,
	type TimelineCache,
	type Table,
} from '../../index';
import type { EditorContext } from '../context';
import { activeTable } from './contextual';
import { target } from './util';

export interface ActiveSlicer {
	frame: SlicerFrameObject;
	slicer: Slicer;
	cache: SlicerCache;
	/** Index in the sheet's drawings. */
	index: number;
}

/** The selected slicer, resolved against the workbook's slicer book. */
export function activeSlicer(ctx: EditorContext): ActiveSlicer | undefined {
	const t = target(ctx);
	const index = ctx.selection.get().drawing;
	if (!t || index === undefined) return undefined;
	const frame = t.ws.drawings[index];
	if (frame?.kind !== 'slicer') return undefined;
	const slicer = t.workbook.slicers?.slicers.find((entry) => entry.name === frame.name);
	const cache = t.workbook.slicers?.caches.find((entry) => entry.name === slicer?.cache);
	return slicer && cache ? { frame, slicer, cache, index } : undefined;
}

/** What Insert > Slicer filters: the table or PivotTable under the active cell. */
export type SlicerSourceAt =
	| { kind: 'table'; table: Table; fields: string[] }
	| { kind: 'pivot'; pivot: PivotTable; fields: string[] };

export function slicerSourceAt(ctx: EditorContext): SlicerSourceAt | undefined {
	const t = target(ctx);
	if (!t) return undefined;
	const pivot = pivotTableAt(t.ws, t.active);
	if (pivot)
		return {
			kind: 'pivot',
			pivot,
			fields: pivot.fields.filter((field) => field !== pivot.pageFilter?.field),
		};
	const table = activeTable(ctx);
	if (table?.headerRow && !table.filterNotModeled)
		return { kind: 'table', table, fields: table.columns.map((column) => column.name) };
	return undefined;
}

export interface ActiveTimeline {
	frame: SlicerFrameObject;
	timeline: Timeline;
	cache: TimelineCache;
	/** Index in the sheet's drawings. */
	index: number;
}

/** The selected timeline, resolved against the workbook's slicer book. */
export function activeTimeline(ctx: EditorContext): ActiveTimeline | undefined {
	const t = target(ctx);
	const index = ctx.selection.get().drawing;
	if (!t || index === undefined) return undefined;
	const frame = t.ws.drawings[index];
	if (frame?.kind !== 'timeline') return undefined;
	const timeline = t.workbook.slicers?.timelines.find((entry) => entry.name === frame.name);
	const cache = t.workbook.slicers?.timelineCaches.find((entry) => entry.name === timeline?.cache);
	return timeline && cache ? { frame, timeline, cache, index } : undefined;
}

/** What Insert > Timeline filters: the PivotTable under the active cell and its date fields. */
export function timelineSourceAt(
	ctx: EditorContext,
): { pivot: PivotTable; fields: string[] } | undefined {
	const t = target(ctx);
	const pivot = t ? pivotTableAt(t.ws, t.active) : undefined;
	return pivot?.dateFields?.length ? { pivot, fields: [...pivot.dateFields] } : undefined;
}
