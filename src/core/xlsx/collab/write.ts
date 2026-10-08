// Model to shared document: writes only what differs, so a local edit becomes a minimal
// transaction and concurrent edits to other cells, rows or sheets are left alone. Call inside
// `doc.transact`.
import type * as Y from 'yjs';
import {
	type CellRange,
	formatAddress,
	normalizeRange,
	parseAddress,
	rangeContains,
} from '../address';
import { forEachCell, forEachCellInRange, getCell } from '../cells';
import type { Workbook, Worksheet } from '../model';
import { encodeCell } from './codec';
import { writeLines } from './lines';
import { writeNames } from './names';
import {
	SCHEMA_VERSION,
	type SharedWorkbook,
	childMap,
	newSheetMap,
	setIfChanged,
	sheetPart,
} from './schema';
import { type SheetEntry, type SheetKeys, reorder, sheetEntries } from './sheets';
import { StyleBridge, sharedStyleKey } from './styles';

/** What a write covers: the whole workbook, or one sheet (optionally only some cell ranges). */
export type WriteScope = { kind: 'all' } | { kind: 'sheet'; sheet: number; ranges?: CellRange[] };

/** Positions a range covers are probed one by one up to this size; larger ranges scan. */
const PROBE_LIMIT = 4096;

export function writeWorkbook(
	shared: SharedWorkbook,
	keys: SheetKeys,
	workbook: Workbook,
	scope: WriteScope,
): void {
	const { meta, sheets } = shared;
	setIfChanged(meta, 'v', SCHEMA_VERSION);
	let defaultKey = meta.get('defaultStyle');
	if (typeof defaultKey !== 'string') {
		const base = workbook.styles[0];
		defaultKey = base ? sharedStyleKey(base) : '';
		meta.set('defaultStyle', defaultKey);
	}
	const styles = new StyleBridge(workbook, shared.styles, defaultKey as string, true);
	styles.keyOf(0);
	const entries = sheetEntries(sheets);
	const known = workbook.sheets.map((sheet) => keys.find(sheet, entries));
	const full = scope.kind === 'all' || known.some((key) => key === undefined || !sheets.has(key));
	if (full) setIfChanged(meta, 'date1904', workbook.date1904);
	const keyList = workbook.sheets.map((sheet) => keys.resolve(sheet, entries, sheets));

	if (full) {
		const live = new Set(keyList);
		for (const key of [...sheets.keys()]) if (!live.has(key)) sheets.delete(key);
	}
	workbook.sheets.forEach((sheet, index) => {
		const key = keyList[index];
		if (key === undefined) return;
		let map = childMap(sheets, key);
		const created = !map;
		if (!map) {
			map = newSheetMap();
			sheets.set(key, map);
		}
		if (!full && !created && scope.kind === 'sheet' && scope.sheet !== index) {
			writeProps(map, sheet);
			return;
		}
		writeProps(map, sheet);
		writeLines(
			sheet,
			{
				rows: sheetPart(map, 'rows'),
				cols: sheetPart(map, 'cols'),
				merges: sheetPart(map, 'merges'),
			},
			styles.keyOf,
		);
		const ranges = !full && scope.kind === 'sheet' ? scope.ranges : undefined;
		writeCells(sheet, sheetPart(map, 'cells'), styles.keyOf, ranges);
	});
	writeOrder(sheets, keyList, entries);
	writeNames(shared.names, workbook.definedNames, (index) => keyList[index]);
}

function writeProps(map: Y.Map<unknown>, sheet: Worksheet): void {
	const props = sheetPart(map, 'props');
	setIfChanged(props, 'name', sheet.name);
	setIfChanged(props, 'sheetId', sheet.sheetId);
	setIfChanged(props, 'state', sheet.state);
	setIfChanged(props, 'tab', sheet.tabColor ? structuredClone(sheet.tabColor) : undefined);
	setIfChanged(props, 'rh', sheet.defaultRowHeight);
	setIfChanged(props, 'cw', sheet.defaultColWidth);
}

function writeOrder(sheets: Y.Map<unknown>, keyList: string[], entries: SheetEntry[]): void {
	const orders = new Map(entries.map((entry) => [entry.key, entry.order]));
	const current = keyList.map((key) => {
		const props = childMap(sheets, key);
		return props && childMap(props, 'props')?.has('order') ? orders.get(key) : undefined;
	});
	for (const [index, order] of reorder(current)) {
		const key = keyList[index];
		const props = key === undefined ? undefined : childMap(sheets, key);
		if (props) sheetPart(props, 'props').set('order', order);
	}
}

function writeCells(
	sheet: Worksheet,
	cells: Y.Map<unknown>,
	keyOf: (styleId: number | undefined) => string | undefined,
	ranges?: CellRange[],
): void {
	const write = (row: number, col: number): string => {
		const key = formatAddress({ row, col });
		const cell = getCell(sheet, row, col);
		setIfChanged(cells, key, cell && encodeCell(cell, keyOf));
		return key;
	};
	if (!ranges) {
		const live = new Set<string>();
		forEachCell(sheet, (_cell, row, col) => live.add(write(row, col)));
		for (const key of [...cells.keys()]) if (!live.has(key)) cells.delete(key);
		return;
	}
	for (const raw of ranges) {
		const range = normalizeRange(raw);
		const area = (range.end.row - range.start.row + 1) * (range.end.col - range.start.col + 1);
		if (area <= PROBE_LIMIT) {
			for (let row = range.start.row; row <= range.end.row; row++)
				for (let col = range.start.col; col <= range.end.col; col++) write(row, col);
			continue;
		}
		const live = new Set<string>();
		forEachCellInRange(sheet, range, (_cell, row, col) => live.add(write(row, col)));
		for (const key of [...cells.keys()]) {
			const at = parseAddress(key);
			if (!live.has(key) && at && rangeContains(range, at)) cells.delete(key);
		}
	}
}
