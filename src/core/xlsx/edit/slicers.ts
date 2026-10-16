import { createGridMetrics } from '../layout/metrics';
import { anchorToPixelBox, pixelBoxToAnchor } from '../layout/anchors';
import type {
	CellValue,
	PivotTable,
	Slicer,
	SlicerCache,
	Table,
	Workbook,
	Worksheet,
} from '../model';
import { pivotItems, pivotValueKey } from '../pivot/data';
import { slicerFrameXml } from '../slicer/frame';
import { slicerItems, slicerTarget } from '../slicer/items';
import { applyTableFilter } from '../slicer/table-filter';
import { type EditContext, sheetAt } from './context';
import { cloneBook, markDirty, removeSlicers } from '../slicer/book';
import { writePivot } from './pivots';

const SLICER_BOX = { w: 144, h: 188 };

function cacheOf(workbook: Workbook, name: string): SlicerCache {
	const cache = workbook.slicers?.caches.find((entry) => entry.name === name);
	if (!cache) throw new Error('The slicer is not available.');
	return cache;
}

/** Display texts and cell values ready to keep a selection on a pivot field. */
function hiddenFor(pivot: PivotTable, field: string, keys: readonly string[] | null): CellValue[] {
	if (keys === null) return [];
	const keep = new Set(keys);
	return pivotItems(pivot, field).filter((value) => !keep.has(pivotValueKey(value)));
}

/**
 * Filters the table column or PivotTable field behind a slicer cache to the given item keys, or
 * clears the filter (`null`). One undoable step; a PivotTable is refreshed with the new items.
 */
export function setSlicerSelection(
	ctx: EditContext,
	cacheName: string,
	keys: readonly string[] | null,
): void {
	const cache = cacheOf(ctx.workbook, cacheName);
	const target = slicerTarget(ctx.workbook, cache);
	if (!target) throw new Error('The slicer has no data source.');
	if (target.kind === 'table') {
		const { sheet, table, offset } = target;
		ctx.run(
			keys === null ? 'Clear slicer filter' : 'Filter slicer',
			'view',
			[{ kind: 'sheet', sheet: target.sheetIndex }],
			() => {
				const column = table.columns[offset]!;
				if (keys === null) delete column.filter;
				else
					column.filter = {
						values: keys.filter((key) => key !== ''),
						...(keys.includes('') ? { blank: true } : {}),
					};
				applyTableFilter(ctx.workbook, sheet, table);
			},
			{ sheet: target.sheetIndex, structural: true },
		);
		return;
	}
	const { pivot } = target;
	const filters = (pivot.itemFilters ?? []).filter((entry) => entry.field !== cache.sourceName);
	const hidden = hiddenFor(pivot, cache.sourceName, keys);
	if (hidden.length) filters.push({ field: cache.sourceName, hidden });
	const next: PivotTable = { ...pivot };
	if (filters.length) next.itemFilters = filters;
	else delete next.itemFilters;
	writePivot(
		ctx,
		target.sheetIndex,
		next,
		pivot,
		keys === null ? 'Clear slicer filter' : 'Filter slicer',
		() => markDirty(ctx.workbook),
	);
}

/** Item keys currently shown as selected, for a slicer cache. */
export const selectedItemKeys = (workbook: Workbook, cacheName: string): string[] =>
	slicerItems(workbook, cacheOf(workbook, cacheName))
		.filter((item) => item.selected)
		.map((item) => item.key);

export interface SlicerPatch {
	caption?: string;
	style?: string;
	columnCount?: number;
	rowHeight?: number;
	showCaption?: boolean;
	sortOrder?: 'ascending' | 'descending';
	hideNoData?: boolean;
}

/** Changes a slicer's caption, style, columns, button height or its cache's sort (one undo step). */
export function updateSlicer(ctx: EditContext, name: string, patch: SlicerPatch): void {
	const slicer = ctx.workbook.slicers?.slicers.find((entry) => entry.name === name);
	if (!slicer) throw new Error('The slicer is not available.');
	if (
		(patch.columnCount !== undefined &&
			(!Number.isInteger(patch.columnCount) ||
				patch.columnCount < 1 ||
				patch.columnCount > 20000)) ||
		(patch.rowHeight !== undefined && !(patch.rowHeight >= 38100 && patch.rowHeight <= 9144000)) ||
		(patch.caption !== undefined && patch.caption.length > 255)
	)
		throw new Error('Choose a valid slicer setting.');
	ctx.run(
		'Format slicer',
		'view',
		[{ kind: 'meta' }],
		() => {
			const book = cloneBook(ctx.workbook);
			const entry = book.slicers.find((item) => item.name === name)!;
			const cache = book.caches.find((item) => item.name === entry.cache)!;
			for (const key of ['caption', 'style', 'rowHeight', 'showCaption'] as const)
				if (patch[key] !== undefined)
					(entry as unknown as Record<string, unknown>)[key] = patch[key];
			if (patch.columnCount !== undefined) {
				if (patch.columnCount === 1) delete entry.columnCount;
				else entry.columnCount = patch.columnCount;
			}
			if (patch.sortOrder !== undefined) cache.sortOrder = patch.sortOrder;
			if (patch.hideNoData !== undefined) {
				if (patch.hideNoData) cache.hideNoData = true;
				else delete cache.hideNoData;
			}
			book.dirty = true;
			ctx.workbook.slicers = book;
		},
		{ structural: false },
	);
}

function frameBounds(sheet: Worksheet): number {
	const metrics = createGridMetrics(sheet, { zoom: 100 });
	let right = 0;
	for (const drawing of sheet.drawings)
		if (drawing.kind === 'slicer' || drawing.kind === 'timeline') {
			const box = anchorToPixelBox(sheet, drawing.anchor, metrics);
			right = Math.max(right, box.x + box.w);
		}
	return right;
}

export type SlicerSourceRef =
	| { kind: 'table'; table: string; column: string }
	| { kind: 'pivot'; pivot: string; field: string };

const identifier = (text: string): string =>
	text.replace(/[^\p{L}\p{N}_]/gu, '_').replace(/^(\d)/, '_$1') || '_';

function unique(base: string, taken: (candidate: string) => boolean): string {
	if (!taken(base)) return base;
	for (let n = 1; ; n++) if (!taken(`${base} ${n}`)) return `${base} ${n}`;
}

/** Inserts a slicer for a table column or PivotTable field on sheet `s` (one undo step). */
export function insertSlicer(ctx: EditContext, s: number, ref: SlicerSourceRef): string {
	const { workbook } = ctx;
	const sheet = sheetAt(workbook, s);
	let table: Table | undefined;
	let pivot: PivotTable | undefined;
	let field: string;
	if (ref.kind === 'table') {
		table = sheet.tables.find(
			(entry) => entry.name === ref.table || entry.displayName === ref.table,
		);
		if (!table?.headerRow || table.filterNotModeled)
			throw new Error('Choose a table with a header row and value filters.');
		field = ref.column;
		if (!table.columns.some((column) => column.name === field))
			throw new Error('Choose a column of the table.');
	} else {
		pivot = sheet.pivotTables?.find((entry) => entry.name === ref.pivot);
		if (!pivot) throw new Error('Choose a PivotTable that can be refreshed.');
		field = ref.field;
		if (!pivot.fields.includes(field) || pivot.pageFilter?.field === field)
			throw new Error('Choose a PivotTable field that is not the report filter.');
	}
	const names = new Set([
		...(workbook.slicers?.slicers ?? []).map((entry) => entry.name.toLowerCase()),
		...(workbook.slicers?.caches ?? []).map((entry) => entry.name.toLowerCase()),
		...workbook.definedNames.map((entry) => entry.name.toLowerCase()),
	]);
	const slicerName = unique(field, (c) => names.has(c.toLowerCase()));
	const cacheName = unique(`Slicer_${identifier(field)}`, (c) =>
		names.has(c.toLowerCase()),
	).replace(/ /g, '_');
	const metrics = createGridMetrics(sheet, { zoom: 100 });
	const anchorCell = table
		? { row: table.range.start.row, col: table.range.end.col + 2 }
		: { row: pivot!.location.start.row, col: pivot!.location.end.col + 2 };
	const x = Math.max(metrics.colLeft(anchorCell.col), frameBounds(sheet) + 10);
	const box = { x, y: metrics.rowTop(anchorCell.row), ...SLICER_BOX };
	const anchor = pixelBoxToAnchor(sheet, box, metrics, 'twoCell');
	const id =
		Math.max(
			1,
			...sheet.drawings.flatMap((d) =>
				'sourceXml' in d && d.sourceXml
					? [...d.sourceXml.matchAll(/<[\w:]*cNvPr\b[^>]*>/g)].flatMap((m) => {
							const id = /\sid="(\d+)"/.exec(m[0])?.[1];
							return id ? [Number(id)] : [];
						})
					: [],
			),
		) + 1;
	ctx.run(
		'Insert slicer',
		'annotations',
		[{ kind: 'sheet', sheet: s }, { kind: 'meta' }],
		() => {
			const book = cloneBook(workbook);
			let source: SlicerCache['source'];
			if (table) {
				if (table.columns.some((column) => column.sourceId === undefined))
					table.columns.forEach((column, index) => (column.sourceId ??= index + 1));
				source = {
					kind: 'table',
					tableId: table.id,
					column: table.columns.find((column) => column.name === field)!.sourceId!,
				};
			} else {
				const target = pivot!;
				target.slicerCacheId ??=
					Math.max(
						1_000_000,
						...workbook.sheets.flatMap((sh) =>
							(sh.pivotTables ?? []).map((p) => p.slicerCacheId ?? 0),
						),
					) + 1;
				target.dirty = true;
				source = {
					kind: 'pivot',
					pivotCacheId: target.slicerCacheId,
					pivots: [{ tabId: sheet.sheetId, name: target.name }],
				};
			}
			book.caches.push({ name: cacheName, sourceName: field, source });
			const slicer: Slicer = { name: slicerName, cache: cacheName, caption: field, sheet: s };
			book.slicers.push(slicer);
			book.dirty = true;
			workbook.slicers = book;
			workbook.definedNames.push({ name: cacheName, formula: '#N/A' });
			sheet.drawings.push({
				kind: 'slicer',
				anchor,
				name: slicerName,
				sourceXml: slicerFrameXml('slicer', slicerName, id, anchor, {
					x: box.x * 9525,
					y: box.y * 9525,
					cx: box.w * 9525,
					cy: box.h * 9525,
				}),
			});
		},
		{ sheet: s },
	);
	return slicerName;
}

/** Deletes one slicer (one undo step); its cache goes with it when no other slicer uses it. */
export function deleteSlicer(ctx: EditContext, name: string): void {
	if (!ctx.workbook.slicers?.slicers.some((slicer) => slicer.name === name))
		throw new Error('The slicer is not available.');
	ctx.run(
		'Delete slicer',
		'annotations',
		[{ kind: 'workbook' }],
		() => removeSlicers(ctx.workbook, (slicer) => slicer.name === name),
		{},
	);
}
