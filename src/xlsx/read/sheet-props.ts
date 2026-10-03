import { parseAddress, parseRange, type CellRange } from '../address.js';
import type {
	AutoFilter,
	Color,
	ColumnInfo,
	PageSetup,
	SheetFormat,
	SheetProtection,
	SheetView,
} from '../model.js';
import type { XmlElement } from '../../xml/index.js';
import { readModernHash } from './password-hash.js';
import { parseColor } from './style-parts.js';
import { att, boolAttr, numAttr, outerXml, xChildren, xFirst, xText } from './xml-util.js';

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

const FORMAT_FLAGS = ['zeroHeight', 'customHeight', 'thickTop', 'thickBottom'] as const;

export function readSheetFormat(format: XmlElement | undefined): {
	defaultRowHeight: number;
	defaultColWidth?: number;
	format?: SheetFormat;
} {
	const result: { defaultRowHeight: number; defaultColWidth?: number; format?: SheetFormat } = {
		defaultRowHeight: numAttr(format, 'defaultRowHeight') ?? 15,
	};
	const width = numAttr(format, 'defaultColWidth');
	if (width !== undefined) result.defaultColWidth = width;
	const flags: SheetFormat = {};
	for (const key of FORMAT_FLAGS) {
		const value = boolAttr(format, key);
		if (value !== undefined) flags[key] = value;
	}
	const base = numAttr(format, 'baseColWidth');
	if (base !== undefined) flags.baseColWidth = base;
	if (Object.keys(flags).length) result.format = flags;
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

/** The default SpreadsheetML declaration xmldom adds to a serialized fragment's first tag. */
const MAIN_NS_DECL =
	/(?<=^<[^>]*) xmlns="http:\/\/schemas\.openxmlformats\.org\/spreadsheetml\/2006\/main"/;

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
		entry.sourceXml = outerXml(column).replace(MAIN_NS_DECL, '');
		columns.push(entry);
	}
	if (columns.length) filter.columns = columns;
	return filter;
}

/** `headerFooter` element names keyed by their `PageSetup` field. */
export const HEADER_FOOTER_TEXT = {
	header: 'oddHeader',
	footer: 'oddFooter',
	evenHeader: 'evenHeader',
	evenFooter: 'evenFooter',
	firstHeader: 'firstHeader',
	firstFooter: 'firstFooter',
} as const;
export const HEADER_FOOTER_FLAGS = [
	'differentOddEven',
	'differentFirst',
	'scaleWithDoc',
	'alignWithMargins',
] as const;
export type HeaderFooterFields = Pick<
	PageSetup,
	keyof typeof HEADER_FOOTER_TEXT | (typeof HEADER_FOOTER_FLAGS)[number]
>;

/** Every header and footer text and the `headerFooter` flags as written in the file. */
export function readHeaderFooter(node: XmlElement): HeaderFooterFields {
	const out: HeaderFooterFields = {};
	for (const key of HEADER_FOOTER_FLAGS) {
		const value = boolAttr(node, key);
		if (value !== undefined) out[key] = value;
	}
	for (const [key, local] of Object.entries(HEADER_FOOTER_TEXT)) {
		const child = xFirst(node, local);
		if (child) out[key as keyof typeof HEADER_FOOTER_TEXT] = xText(child);
	}
	return out;
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
	if (headerFooter) Object.assign(page, readHeaderFooter(headerFooter));
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
	const modernHash = readModernHash(node);
	if (modernHash) protection.modernHash = modernHash;
	const allow: string[] = [];
	for (const attribute of Array.from(node.attributes)) {
		if (PROTECTION_META.has(attribute.name)) continue;
		if (attribute.value === '0' || attribute.value === 'false') allow.push(attribute.name);
	}
	if (allow.length) protection.allow = allow.sort();
	return protection;
}
