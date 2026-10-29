// View > Custom Views: save the current look of the workbook under a name, show it again, delete
// it. A view records, per sheet, the zoom, grid lines, headings, panes and selection; with
// "hidden rows, columns and filter settings" the hidden rows and columns and the auto-filter; with
// "print settings" the page setup and print options. Hidden rows, columns, the filter range and
// the print area go to hidden `Z_<GUID>_.wvu.*` names, as Excel stores them.
import { columnIndex, columnLabel, parseRange, quoteSheetName } from '../address';
import type {
	AutoFilter,
	CustomSheetView,
	CustomWorkbookView,
	DefinedName,
	Workbook,
	Worksheet,
} from '../model';
import { clearFilteredRows } from '../filter-rows';
import { editColumns } from './columns';
import type { EditContext } from './context';
import { applyFilter } from './filter';

export type CustomViewProblem = 'name-empty' | 'name-too-long' | 'tables' | 'missing';

const MESSAGES: Record<CustomViewProblem, string> = {
	'name-empty': 'Enter a name for the view.',
	'name-too-long': 'A view name can have at most 255 characters.',
	tables: 'Custom views cannot be used in a workbook that contains tables.',
	missing: 'The view does not exist.',
};

export const customViewMessage = (problem: CustomViewProblem): string => MESSAGES[problem];

class CustomViewError extends Error {
	constructor(readonly problem: CustomViewProblem) {
		super(MESSAGES[problem]);
	}
}

export interface CustomViewDraft {
	name: string;
	/** Page setup and print options are part of the view (default true). */
	includePrintSettings?: boolean;
	/** Hidden rows, columns and filter settings are part of the view (default true). */
	includeHiddenRowCol?: boolean;
}

/** The views of the workbook, in the order Excel lists them. */
export const customViewsOf = (workbook: Workbook): CustomWorkbookView[] =>
	workbook.customViews ?? [];

const newGuid = (): string => {
	const id =
		globalThis.crypto?.randomUUID?.() ??
		'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
			const r = Math.floor(Math.random() * 16);
			return (c === 'x' ? r : (r & 3) | 8).toString(16);
		});
	return `{${id.toUpperCase()}}`;
};

/** The prefix of the hidden names that belong to a view (`Z_9EDFE334_..._.wvu.`). */
export const viewNamePrefix = (guid: string): string =>
	`Z_${guid.replace(/[{}]/g, '').replace(/-/g, '_').toUpperCase()}_.wvu.`;

const viewName = (workbook: Workbook, guid: string, suffix: string, sheet: number) =>
	workbook.definedNames.find(
		(n) =>
			n.localSheet === sheet &&
			n.name.toUpperCase() === `${viewNamePrefix(guid)}${suffix}`.toUpperCase(),
	);

/** `Sheet!$3:$3,Sheet!$5:$6` style list of the runs in `indices`. */
function span(indices: number[], prefix: string, label: (n: number) => string): string[] {
	const out: string[] = [];
	const sorted = [...indices].sort((a, b) => a - b);
	for (let i = 0; i < sorted.length;) {
		let j = i;
		while (sorted[j + 1] === sorted[j]! + 1) j++;
		out.push(`${prefix}!${label(sorted[i]!)}:${label(sorted[j]!)}`);
		i = j + 1;
	}
	return out;
}

const manualHiddenRows = (sheet: Worksheet): number[] =>
	[...sheet.rowInfo]
		.filter(([, info]) => info.hidden && (!info.filteredOut || info.manuallyHidden))
		.map(([row]) => row);

const hiddenColumns = (sheet: Worksheet): number[] =>
	sheet.columns.flatMap((info) =>
		info.hidden ? Array.from({ length: info.max - info.min + 1 }, (_v, i) => info.min + i) : [],
	);

/** Row numbers or column letters listed by a `.wvu.Rows` / `.wvu.Cols` name. */
function listed(formula: string, axis: 'row' | 'col'): number[] {
	const out: number[] = [];
	for (const part of formula.split(',')) {
		const ref = part.slice(part.lastIndexOf('!') + 1).replace(/\$/g, '');
		const match = (axis === 'row' ? /^(\d+):(\d+)$/ : /^([A-Za-z]+):([A-Za-z]+)$/).exec(ref);
		if (!match) continue;
		const from = axis === 'row' ? Number(match[1]) - 1 : columnIndex(match[1]!.toUpperCase());
		const to = axis === 'row' ? Number(match[2]) - 1 : columnIndex(match[2]!.toUpperCase());
		for (let n = Math.min(from, to); n <= Math.max(from, to) && out.length < 20000; n++)
			out.push(n);
	}
	return out;
}

const dropViewNames = (workbook: Workbook, guid: string): void => {
	const prefix = viewNamePrefix(guid).toUpperCase();
	workbook.definedNames = workbook.definedNames.filter(
		(name) => !name.name.toUpperCase().startsWith(prefix),
	);
};

function absolute(sheet: Worksheet, range: NonNullable<AutoFilter['range']>): string {
	const cell = (row: number, col: number) => `$${columnLabel(col)}$${row + 1}`;
	return `${quoteSheetName(sheet.name)}!${cell(range.start.row, range.start.col)}:${cell(range.end.row, range.end.col)}`;
}

function recordSheet(
	workbook: Workbook,
	sheet: Worksheet,
	index: number,
	guid: string,
	draft: Required<CustomViewDraft>,
): CustomSheetView {
	const view: CustomSheetView = { guid, view: structuredClone(sheet.view), attrs: {}, extra: [] };
	const add = (suffix: string, formula: string) => {
		const name: DefinedName = {
			name: `${viewNamePrefix(guid)}${suffix}`,
			formula,
			localSheet: index,
			hidden: true,
		};
		workbook.definedNames.push(name);
	};
	if (draft.includeHiddenRowCol) {
		const prefix = quoteSheetName(sheet.name);
		const rows = manualHiddenRows(sheet);
		if (rows.length) {
			view.hiddenRows = true;
			add('Rows', span(rows, prefix, (n) => `$${n + 1}`).join(','));
		}
		const cols = hiddenColumns(sheet);
		if (cols.length) {
			view.hiddenColumns = true;
			add('Cols', span(cols, prefix, (n) => `$${columnLabel(n)}`).join(','));
		}
		if (sheet.autoFilter) {
			view.filter = structuredClone(sheet.autoFilter);
			view.showAutoFilter = true;
			add('FilterData', absolute(sheet, sheet.autoFilter.range));
		}
	}
	if (draft.includePrintSettings) {
		if (sheet.pageSetup) {
			view.pageSetup = structuredClone(sheet.pageSetup);
			if (sheet.pageSetup.printArea) add('PrintArea', absolute(sheet, sheet.pageSetup.printArea));
		}
		if (sheet.printOptions) view.printOptions = structuredClone(sheet.printOptions);
	}
	return view;
}

/** Saves the current look of every sheet as a view; a view of the same name is replaced. */
export function addCustomView(ctx: EditContext, draft: CustomViewDraft): string {
	const { workbook } = ctx;
	const name = draft.name.trim();
	if (!name) throw new CustomViewError('name-empty');
	if (name.length > 255) throw new CustomViewError('name-too-long');
	if (workbook.sheets.some((sheet) => sheet.tables.length)) throw new CustomViewError('tables');
	const options = {
		name,
		includePrintSettings: draft.includePrintSettings ?? true,
		includeHiddenRowCol: draft.includeHiddenRowCol ?? true,
	};
	return ctx.run('Add custom view', 'view', [{ kind: 'workbook' }], () => {
		const same = customViewsOf(workbook).find((v) => v.name.toLowerCase() === name.toLowerCase());
		if (same) removeView(workbook, same.guid);
		const guid = newGuid();
		const active = workbook.sheets[workbook.activeSheet];
		const attrs: Record<string, string> = active ? { activeSheetId: String(active.sheetId) } : {};
		workbook.customViews = [
			...customViewsOf(workbook),
			{
				guid,
				name,
				includePrintSettings: options.includePrintSettings,
				includeHiddenRowCol: options.includeHiddenRowCol,
				attrs,
			},
		];
		workbook.sheets.forEach((sheet, index) => {
			sheet.customViews = [
				...(sheet.customViews ?? []),
				recordSheet(workbook, sheet, index, guid, options),
			];
		});
		return guid;
	});
}

function removeView(workbook: Workbook, guid: string): void {
	workbook.customViews = customViewsOf(workbook).filter((v) => v.guid !== guid);
	for (const sheet of workbook.sheets) {
		const rest = (sheet.customViews ?? []).filter((v) => v.guid !== guid);
		if (rest.length) sheet.customViews = rest;
		else delete sheet.customViews;
	}
	dropViewNames(workbook, guid);
}

export function deleteCustomView(ctx: EditContext, guid: string): void {
	const { workbook } = ctx;
	if (!customViewsOf(workbook).some((v) => v.guid === guid)) throw new CustomViewError('missing');
	ctx.run('Delete custom view', 'view', [{ kind: 'workbook' }], () => removeView(workbook, guid));
}

/** Applies a saved view: every sheet that recorded it takes its look back. */
export function showCustomView(ctx: EditContext, guid: string): void {
	const { workbook } = ctx;
	const book = customViewsOf(workbook).find((v) => v.guid === guid);
	if (!book) throw new CustomViewError('missing');
	ctx.run(
		'Show custom view',
		'view',
		[{ kind: 'workbook' }],
		() => {
			workbook.sheets.forEach((sheet, index) => {
				const saved = sheet.customViews?.find((v) => v.guid === guid);
				if (!saved) return;
				sheet.view = structuredClone(saved.view);
				if (book.includeHiddenRowCol) applyHidden(ctx, workbook, sheet, index, saved);
				if (book.includePrintSettings) applyPrint(workbook, sheet, index, saved);
			});
			const activeId = Number(book.attrs.activeSheetId);
			const active = workbook.sheets.findIndex((s) => s.sheetId === activeId);
			if (active >= 0) workbook.activeSheet = active;
		},
		{ structural: true },
	);
}

function applyHidden(
	ctx: EditContext,
	workbook: Workbook,
	sheet: Worksheet,
	index: number,
	saved: CustomSheetView,
): void {
	// Rows and columns hidden by hand are replaced by the view's; filtered rows follow the filter.
	for (const row of manualHiddenRows(sheet)) {
		const info = { ...sheet.rowInfo.get(row) };
		if (info.filteredOut) delete info.manuallyHidden;
		else delete info.hidden;
		if (Object.keys(info).length) sheet.rowInfo.set(row, info);
		else sheet.rowInfo.delete(row);
	}
	const hide = viewName(workbook, saved.guid, 'Rows', index);
	for (const row of hide ? listed(hide.formula, 'row') : []) {
		const info = { ...sheet.rowInfo.get(row) };
		info.hidden = true;
		sheet.rowInfo.set(row, info);
	}
	const open = hiddenColumns(sheet);
	if (open.length)
		editColumns(sheet, open, (info) => {
			delete info.hidden;
		});
	const cols = viewName(workbook, saved.guid, 'Cols', index);
	const toHide = cols ? listed(cols.formula, 'col') : [];
	if (toHide.length)
		editColumns(sheet, toHide, (info) => {
			info.hidden = true;
		});
	clearFilteredRows(sheet);
	if (saved.filter) {
		sheet.autoFilter = structuredClone(saved.filter);
		applyFilter(ctx, sheet, sheet.autoFilter);
	} else delete sheet.autoFilter;
}

function applyPrint(
	workbook: Workbook,
	sheet: Worksheet,
	index: number,
	saved: CustomSheetView,
): void {
	if (saved.pageSetup) sheet.pageSetup = structuredClone(saved.pageSetup);
	else delete sheet.pageSetup;
	if (saved.printOptions) sheet.printOptions = structuredClone(saved.printOptions);
	else delete sheet.printOptions;
	const area = viewName(workbook, saved.guid, 'PrintArea', index);
	const range = area
		? parseRange(area.formula.slice(area.formula.lastIndexOf('!') + 1).replace(/\$/g, ''))
		: undefined;
	if (range) sheet.pageSetup = { ...sheet.pageSetup, printArea: range };
}
