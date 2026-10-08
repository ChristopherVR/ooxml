// Shared document to model, in place: only what a transaction touched is re-read, the sheet
// objects the views hold stay the same, and everything the shared schema does not carry (charts,
// drawings, comments, conditional formats, validations, tables, views) is left as it was.
import { type CellRange, formatAddress, parseAddress } from '../address';
import { deleteCell, forEachCell, getCell, putCell } from '../cells';
import type { ExternalChange } from '../edit/external';
import { isSpilledCell } from '../formula/spill';
import type { Workbook, Worksheet } from '../model';
import { createWorksheet } from '../workbook';
import { decodeCell, encodeCell } from './codec';
import type { Dirty } from './dirty';
import { applyColumns, applyMerges, applyRows } from './lines';
import { readNames } from './names';
import { type SharedWorkbook, childMap, jsonEqual } from './schema';
import { type SheetEntry, type SheetKeys, sheetEntries } from './sheets';
import { StyleBridge } from './styles';

type Touched = { sheet: number; row: number; col: number };

/** Applies `dirty` to the workbook; undefined when nothing in the model changed. */
export function applyShared(
	shared: SharedWorkbook,
	keys: SheetKeys,
	workbook: Workbook,
	dirty: Dirty,
	label = 'Remote change',
): ExternalChange | undefined {
	const defaultKey = shared.meta.get('defaultStyle');
	const styles = new StyleBridge(workbook, shared.styles, String(defaultKey ?? ''), false);
	const entries = sheetEntries(shared.sheets);
	let structural = false;
	if (dirty.all || dirty.list) structural = applyList(workbook, keys, entries);
	const date1904 = shared.meta.get('date1904');
	if (dirty.all && typeof date1904 === 'boolean') workbook.date1904 = date1904;

	const indexOf = new Map<string, number>();
	workbook.sheets.forEach((sheet, index) => {
		const key = keys.find(sheet, entries);
		if (key !== undefined) indexOf.set(key, index);
	});
	const touched: Touched[] = [];
	const ranges: CellRange[] = [];
	const sheetsTouched = new Set<number>();
	const visit = dirty.all ? entries.map((e) => e.key) : [...dirty.sheets.keys()];
	for (const key of visit) {
		const index = indexOf.get(key);
		const sheet = index === undefined ? undefined : workbook.sheets[index];
		const map = childMap(shared.sheets, key);
		if (!sheet || index === undefined || !map) continue;
		const part = dirty.sheets.get(key);
		const full = dirty.all || !part || part.full;
		const rows = childMap(map, 'rows');
		if (rows && (full || part.rows.size)) {
			applyRows(sheet, rows, styles.idOf, full ? undefined : part.rows);
			structural = true;
		}
		const cols = childMap(map, 'cols');
		if (cols && (full || part.cols)) {
			applyColumns(sheet, cols, styles.idOf);
			structural = true;
		}
		const merges = childMap(map, 'merges');
		if (merges && (full || part.merges)) {
			applyMerges(sheet, merges);
			sheetsTouched.add(index);
		}
		const cells = childMap(map, 'cells');
		if (!cells) continue;
		const put = (address: string): void => {
			const at = parseAddress(address);
			if (!at) return;
			const raw = cells.get(address);
			const local = getCell(sheet, at.row, at.col);
			if (local && jsonEqual(encodeCell(local, styles.keyOf), raw)) return;
			const cell = raw === undefined ? undefined : decodeCell(raw, styles.idOf);
			if (cell) putCell(sheet, at.row, at.col, cell);
			else if (local && !isSpilledCell(local)) deleteCell(sheet, at.row, at.col);
			else return;
			touched.push({ sheet: index, row: at.row, col: at.col });
			ranges.push({ start: at, end: { ...at } });
			sheetsTouched.add(index);
		};
		if (full) {
			const stale: string[] = [];
			forEachCell(sheet, (cell, row, col) => {
				const address = formatAddress({ row, col });
				if (!isSpilledCell(cell) && !cells.has(address)) stale.push(address);
			});
			for (const address of [...cells.keys(), ...stale]) put(address);
		} else for (const address of part.cells) put(address);
	}
	let namesChanged = false;
	if (dirty.all || dirty.list || dirty.names) {
		const names = readNames(shared.names, (key) => indexOf.get(key));
		namesChanged = !jsonEqual(names, workbook.definedNames);
		if (namesChanged) workbook.definedNames = names;
	}
	if (!structural && !sheetsTouched.size && !namesChanged) return undefined;
	const change: ExternalChange = { label, structural, cells: touched };
	const [only] = sheetsTouched;
	if (sheetsTouched.size === 1 && only !== undefined) change.sheet = only;
	if (!structural && ranges.length && ranges.length <= 1024) change.ranges = ranges;
	return change;
}

/** Rebuilds the sheet list from the shared entries, keeping known sheet objects. */
function applyList(workbook: Workbook, keys: SheetKeys, entries: SheetEntry[]): boolean {
	if (!entries.length) return false; // A workbook keeps at least one sheet.
	const byKey = new Map<string, Worksheet>();
	for (const sheet of workbook.sheets) {
		const key = keys.find(sheet, entries);
		if (key !== undefined && !byKey.has(key)) byKey.set(key, sheet);
	}
	const active = workbook.sheets[workbook.activeSheet];
	let changed = entries.length !== workbook.sheets.length;
	const next = entries.map((entry, index) => {
		const sheet = byKey.get(entry.key) ?? createWorksheet(entry.name, entry.sheetId);
		keys.remember(sheet, entry.key);
		if (
			workbook.sheets[index] !== sheet ||
			sheet.name !== entry.name ||
			sheet.state !== entry.state ||
			sheet.sheetId !== entry.sheetId
		)
			changed = true;
		sheet.name = entry.name;
		sheet.sheetId = entry.sheetId;
		sheet.state = entry.state;
		if (entry.tabColor) sheet.tabColor = entry.tabColor;
		else delete sheet.tabColor;
		if (entry.defaultRowHeight !== undefined) sheet.defaultRowHeight = entry.defaultRowHeight;
		if (entry.defaultColWidth !== undefined) sheet.defaultColWidth = entry.defaultColWidth;
		else delete sheet.defaultColWidth;
		return sheet;
	});
	workbook.sheets.splice(0, workbook.sheets.length, ...next);
	const kept = active ? next.indexOf(active) : -1;
	workbook.activeSheet = kept >= 0 ? kept : Math.min(workbook.activeSheet, next.length - 1);
	return changed;
}
