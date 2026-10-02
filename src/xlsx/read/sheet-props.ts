import { parseAddress, parseRange, type CellRange } from '../address.js';
import type {
	AutoFilter,
	Color,
	ColumnInfo,
	PageSetup,
	SheetProtection,
	SheetView,
} from '../model.js';
import type { XmlElement } from '../../xml/index.js';
import { parseColor } from './style-parts.js';
import { att, boolAttr, numAttr, xChildren, xFirst, xText } from './xml-util.js';

export const defaultSheetView = (): SheetView => ({
	showGridLines: true,
	showHeaders: true,
	showZeros: true,
	rightToLeft: false,
	zoom: 100,
});

/** A space-separated `sqref` list. */
export function parseSqref(sqref: string | undefined): CellRange[] {
	if (!sqref) return [];
	return sqref
		.split(/\s+/)
		.filter(Boolean)
		.map((part) => parseRange(part))
		.filter((range): range is CellRange => range !== undefined);
}

export const readTabColor = (
	sheetPr: XmlElement | undefined,
	palette?: readonly string[],
): Color | undefined => parseColor(xFirst(sheetPr, 'tabColor'), palette);

export const readFitToPage = (sheetPr: XmlElement | undefined): boolean =>
	boolAttr(xFirst(sheetPr, 'pageSetUpPr'), 'fitToPage') ?? false;

/** The first `sheetView` (the one Excel shows), with frozen panes and the active selection. */
export function readSheetView(sheetViews: XmlElement | undefined): SheetView {
	const view = defaultSheetView();
	const node = xFirst(sheetViews, 'sheetView');
	if (!node) return view;
	view.showGridLines = boolAttr(node, 'showGridLines', true);
	view.showHeaders = boolAttr(node, 'showRowColHeaders', true);
	view.showZeros = boolAttr(node, 'showZeros', true);
	view.rightToLeft = boolAttr(node, 'rightToLeft', false);
	view.zoom = numAttr(node, 'zoomScale') || 100;
	if (boolAttr(node, 'showFormulas')) view.showFormulas = true;
	const pane = xFirst(node, 'pane');
	const state = att(pane, 'state');
	const frozen = pane && (state === 'frozen' || state === 'frozenSplit');
	if (frozen) {
		const rows = Math.max(0, Math.round(numAttr(pane, 'ySplit') ?? 0));
		const cols = Math.max(0, Math.round(numAttr(pane, 'xSplit') ?? 0));
		if (rows || cols) view.freeze = { rows, cols };
	}
	const topLeft = parseAddress(
		(frozen ? att(pane, 'topLeftCell') : att(node, 'topLeftCell')) ?? '',
	);
	if (topLeft) view.topLeft = topLeft;
	const selections = xChildren(node, 'selection');
	const activePane = att(pane, 'activePane') ?? 'topLeft';
	const selection =
		selections.find((sel) => (att(sel, 'pane') ?? 'topLeft') === activePane) ?? selections[0];
	if (selection) {
		const ranges = parseSqref(att(selection, 'sqref') ?? att(selection, 'activeCell') ?? 'A1');
		const active = parseAddress(att(selection, 'activeCell') ?? '') ??
			ranges[0]?.start ?? { row: 0, col: 0 };
		view.selection = {
			active,
			ranges: ranges.length ? ranges : [{ start: active, end: { ...active } }],
		};
	}
	return view;
}

export function readSheetFormat(format: XmlElement | undefined): {
	defaultRowHeight: number;
	defaultColWidth?: number;
} {
	const result: { defaultRowHeight: number; defaultColWidth?: number } = {
		defaultRowHeight: numAttr(format, 'defaultRowHeight') ?? 15,
	};
	const width = numAttr(format, 'defaultColWidth');
	if (width !== undefined) result.defaultColWidth = width;
	return result;
}

export function readColumns(cols: XmlElement | undefined, xfMap: readonly number[]): ColumnInfo[] {
	const out: ColumnInfo[] = [];
	for (const col of cols ? xChildren(cols, 'col') : []) {
		const min = (numAttr(col, 'min') ?? 1) - 1;
		const max = (numAttr(col, 'max') ?? min + 1) - 1;
		const info: ColumnInfo = { min, max };
		const width = numAttr(col, 'width');
		if (width !== undefined) info.width = width;
		if (boolAttr(col, 'customWidth')) info.customWidth = true;
		if (boolAttr(col, 'hidden')) info.hidden = true;
		const style = numAttr(col, 'style');
		const styleId = style === undefined ? 0 : (xfMap[style] ?? 0);
		if (styleId) info.styleId = styleId;
		const outline = numAttr(col, 'outlineLevel');
		if (outline) info.outlineLevel = outline;
		if (boolAttr(col, 'collapsed')) info.collapsed = true;
		if (boolAttr(col, 'bestFit')) info.bestFit = true;
		out.push(info);
	}
	return out;
}

export function readAutoFilter(node: XmlElement | undefined): AutoFilter | undefined {
	const range = parseRange(att(node, 'ref') ?? '');
	if (!node || !range) return undefined;
	const filter: AutoFilter = { range };
	const columns: NonNullable<AutoFilter['columns']> = [];
	for (const column of xChildren(node, 'filterColumn')) {
		const entry: NonNullable<AutoFilter['columns']>[number] = {
			offset: numAttr(column, 'colId') ?? 0,
		};
		const filters = xFirst(column, 'filters');
		if (filters) {
			entry.values = xChildren(filters, 'filter').map((f) => att(f, 'val') ?? '');
			if (boolAttr(filters, 'blank')) entry.blank = true;
		}
		columns.push(entry);
	}
	if (columns.length) filter.columns = columns;
	return filter;
}

const ORIENTATIONS = new Set(['portrait', 'landscape']);

/** Page margins, setup and the odd-page header/footer. */
export function readPageSetup(
	margins: XmlElement | undefined,
	setup: XmlElement | undefined,
	headerFooter: XmlElement | undefined,
	fitToPage: boolean,
): PageSetup | undefined {
	const page: PageSetup = {};
	if (setup) {
		const orientation = att(setup, 'orientation');
		if (orientation && ORIENTATIONS.has(orientation))
			page.orientation = orientation as 'portrait' | 'landscape';
		const paper = numAttr(setup, 'paperSize');
		if (paper !== undefined) page.paperSize = paper;
		const scale = numAttr(setup, 'scale');
		if (scale !== undefined) page.scale = scale;
	}
	if (fitToPage) {
		page.fitToWidth = numAttr(setup, 'fitToWidth') ?? 1;
		page.fitToHeight = numAttr(setup, 'fitToHeight') ?? 1;
	}
	if (margins) {
		const m = (name: string, fallback: number) => numAttr(margins, name) ?? fallback;
		page.margins = {
			left: m('left', 0.7),
			right: m('right', 0.7),
			top: m('top', 0.75),
			bottom: m('bottom', 0.75),
			header: m('header', 0.3),
			footer: m('footer', 0.3),
		};
	}
	const header = xFirst(headerFooter, 'oddHeader');
	if (header) page.header = xText(header);
	const footer = xFirst(headerFooter, 'oddFooter');
	if (footer) page.footer = xText(footer);
	return Object.keys(page).length ? page : undefined;
}

/** Attributes of `sheetProtection` that are hash material, not "allowed action" flags. */
const PROTECTION_META = new Set([
	'sheet',
	'password',
	'algorithmName',
	'hashValue',
	'saltValue',
	'spinCount',
]);

/** Reads sheet protection; `allow` lists the actions whose protection flag is off (`="0"`). */
export function readProtection(node: XmlElement | undefined): SheetProtection | undefined {
	if (!node) return undefined;
	const protection: SheetProtection = { sheet: boolAttr(node, 'sheet', false) };
	const password = att(node, 'password');
	if (password) protection.passwordHash = password;
	const allow: string[] = [];
	for (const attribute of Array.from(node.attributes)) {
		if (PROTECTION_META.has(attribute.name)) continue;
		if (attribute.value === '0' || attribute.value === 'false') allow.push(attribute.name);
	}
	if (allow.length) protection.allow = allow.sort();
	return protection;
}
