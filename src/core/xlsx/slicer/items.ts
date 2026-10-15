import { getCell } from '../cells';
import { compareCellValues } from '../edit/sort';
import type { CellValue, PivotTable, SlicerCache, Table, Workbook, Worksheet } from '../model';
import { pivotItemLabel, pivotItems, pivotValueKey } from '../pivot/data';
import { hiddenItemKeys, itemsWithData } from '../pivot/item-filters';
import { rowMatchesFilter } from '../filter-rows';
import { displayText } from '../display-text';
import { tableAutoFilter, tableFilterBounds } from './table-filter';

/** Excel's four slicer button states. */
export type SlicerItemState = 'selected' | 'unselected' | 'selectedNoData' | 'unselectedNoData';

export interface SlicerItem {
	/** Stable identity: the display text (tables) or the pivot item key. */
	key: string;
	label: string;
	selected: boolean;
	hasData: boolean;
	state: SlicerItemState;
}

/** A slicer cache's resolved target. */
export type SlicerTarget =
	| { kind: 'table'; sheet: Worksheet; sheetIndex: number; table: Table; offset: number }
	| { kind: 'pivot'; sheet: Worksheet; sheetIndex: number; pivot: PivotTable };

export function slicerTarget(workbook: Workbook, cache: SlicerCache): SlicerTarget | undefined {
	const source = cache.source;
	for (const [sheetIndex, sheet] of workbook.sheets.entries()) {
		if (source.kind === 'table') {
			const table = sheet.tables.find((entry) => entry.id === source.tableId);
			const offset = table?.columns.findIndex((column) => column.sourceId === source.column) ?? -1;
			if (table && offset >= 0) return { kind: 'table', sheet, sheetIndex, table, offset };
		} else {
			const first = source.pivots[0];
			const pivot =
				first && sheet.sheetId === first.tabId
					? sheet.pivotTables?.find((entry) => entry.name === first.name)
					: undefined;
			if (pivot) return { kind: 'pivot', sheet, sheetIndex, pivot };
		}
	}
	return undefined;
}

const stateOf = (selected: boolean, hasData: boolean): SlicerItemState =>
	selected
		? hasData
			? 'selected'
			: 'selectedNoData'
		: hasData
			? 'unselected'
			: 'unselectedNoData';

function tableItems(workbook: Workbook, target: Extract<SlicerTarget, { kind: 'table' }>) {
	const { sheet, table, offset } = target;
	const bounds = tableFilterBounds(table);
	const column = table.columns[offset]!;
	const col = bounds.start.col + offset;
	const filter = tableAutoFilter(table);
	const others = filter && {
		...filter,
		columns: (filter.columns ?? []).filter((entry) => entry.offset !== offset),
	};
	const seen = new Map<string, { value: CellValue; hasData: boolean }>();
	for (let row = bounds.start.row + 1; row <= bounds.end.row; row++) {
		const cell = getCell(sheet, row, col);
		const text = displayText(workbook, cell);
		const visible = !others?.columns?.length || rowMatchesFilter(workbook, sheet, row, others);
		const known = seen.get(text);
		if (known) known.hasData ||= visible;
		else seen.set(text, { value: cell?.value ?? null, hasData: visible });
	}
	const chosen = column.filter;
	const entries = [...seen.entries()].sort(([, a], [, b]) => compareCellValues(a.value, b.value));
	return entries.map(([text, info]): SlicerItem => {
		const selected = chosen ? (text === '' ? !!chosen.blank : chosen.values.includes(text)) : true;
		return {
			key: text,
			label: text === '' ? '(blank)' : text,
			selected,
			hasData: info.hasData,
			state: stateOf(selected, info.hasData),
		};
	});
}

function pivotSlicerItems(pivot: PivotTable, field: string): SlicerItem[] {
	const hidden = hiddenItemKeys(pivot, field);
	const withData = itemsWithData(pivot, field);
	return pivotItems(pivot, field).map((value): SlicerItem => {
		const key = pivotValueKey(value);
		const selected = !hidden.has(key);
		const hasData = withData.has(key);
		return {
			key,
			label: String(pivotItemLabel(value)),
			selected,
			hasData,
			state: stateOf(selected, hasData),
		};
	});
}

/** The slicer's buttons in display order (sorted, no-data items placed per the cache). */
export function slicerItems(workbook: Workbook, cache: SlicerCache): SlicerItem[] {
	const target = slicerTarget(workbook, cache);
	if (!target) return [];
	let items =
		target.kind === 'table'
			? tableItems(workbook, target)
			: pivotSlicerItems(target.pivot, cache.sourceName);
	if (cache.sortOrder === 'descending') items = [...items].reverse();
	if (cache.hideNoData) items = items.filter((item) => item.hasData);
	else if (cache.crossFilter === 'showItemsWithDataAtTop')
		items = [...items.filter((i) => i.hasData), ...items.filter((i) => !i.hasData)];
	return items;
}

/** Whether any item is deselected, which is when Clear Filter is available. */
export const slicerFiltered = (items: readonly SlicerItem[]): boolean =>
	items.some((item) => !item.selected);
