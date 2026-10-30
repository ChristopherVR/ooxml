import { elements, type XmlElement } from '../../xml/index';
import { formatAddress, formatRange, rangeContains, type CellAddress } from '../address';
import { usedRange } from '../cells';
import type { ColumnInfo, Worksheet } from '../model';
import { splitActivePane } from '../layout/split';
import {
	defaultSheetView,
	readAutoFilter,
	readFitToPage,
	readProtection,
	readSheetFormat,
	readSheetView,
	readTabColor,
} from '../read/sheet-props';
import { outerXml } from '../read/xml-util';
import { filterColumnXml, unresolvedFilter } from './filter-columns';
import { sortStateXml } from './sort-state';
import { mergedAttrs } from './attr-merge';
import { modernHashValues } from './password-hash';
import { sameModel, snapshotElement, snapshotXml } from './snapshot';
import type { StyleWriter } from './styles';
import { colorXml } from './style-xml';
import { attrs, el, escapeAttr } from './xml-out';

const sqref = (ranges: readonly { start: CellAddress; end: CellAddress }[]) =>
	ranges.map((range) => formatRange(range)).join(' ');

export function sheetPrXml(sheet: Worksheet): string {
	const source = snapshotElement(sheet, 'sheetPr');
	const fit =
		sheet.pageSetup?.fitToWidth !== undefined || sheet.pageSetup?.fitToHeight !== undefined;
	if (source && sameModel(readTabColor(source), sheet.tabColor) && readFitToPage(source) === fit)
		return snapshotXml(sheet, 'sheetPr') ?? '';
	const sourceAttrs = source
		? Array.from(source.attributes)
				.filter((a) => !a.name.startsWith('xmlns'))
				.map((a) => ` ${a.name}="${escapeAttr(a.value)}"`)
				.join('')
		: '';
	const children = source ? elements(source) : [];
	const outline = children.find((node) => node.localName === 'outlinePr');
	const setUp = children.find((node) => node.localName === 'pageSetUpPr');
	let setUpXml = '';
	if (setUp || fit) {
		const keep = setUp
			? Array.from(setUp.attributes)
					.filter((a) => a.name !== 'fitToPage' && !a.name.startsWith('xmlns'))
					.map((a) => ` ${a.name}="${escapeAttr(a.value)}"`)
					.join('')
			: '';
		setUpXml = `<pageSetUpPr${keep}${fit ? ' fitToPage="1"' : ''}/>`;
		if (setUpXml === '<pageSetUpPr/>') setUpXml = '';
	}
	const inner =
		colorXml(sheet.tabColor, 'tabColor') + (outline ? outerXml(outline) : '') + setUpXml;
	if (!inner && !sourceAttrs) return '';
	return `<sheetPr${sourceAttrs}>${inner}</sheetPr>`;
}

export function dimensionXml(sheet: Worksheet): string {
	const used = usedRange(sheet);
	return `<dimension ref="${used ? formatRange(used) : 'A1'}"/>`;
}

function selectionXml(
	selection: NonNullable<Worksheet['view']['selection']>,
	pane: string | undefined,
): string {
	const activeCellId = selection.ranges.findIndex((range) =>
		rangeContains(range, selection.active),
	);
	return el('selection', {
		pane,
		activeCell: formatAddress(selection.active),
		activeCellId: activeCellId > 0 ? activeCellId : undefined,
		sqref: sqref(selection.ranges) || formatAddress(selection.active),
	});
}

export function sheetViewsXml(sheet: Worksheet, tabSelected: boolean): string {
	const source = snapshotElement(sheet, 'sheetViews');
	if (source && sameModel(readSheetView(source), sheet.view)) {
		const first = elements(source).find((node) => node.localName === 'sheetView');
		if (first) {
			if (tabSelected) first.setAttribute('tabSelected', '1');
			else first.removeAttribute('tabSelected');
		}
		return outerXml(source);
	}
	// Workbooks built in code may omit the view; write the defaults then.
	const view = sheet.view ?? defaultSheetView();
	let inner = '';
	let pane: string | undefined;
	if (view.freeze && (view.freeze.rows > 0 || view.freeze.cols > 0)) {
		const { rows, cols } = view.freeze;
		pane = rows && cols ? 'bottomRight' : rows ? 'bottomLeft' : 'topRight';
		inner += el('pane', {
			xSplit: cols || undefined,
			ySplit: rows || undefined,
			// Excel rejects a scrolled pane whose top-left cell lies inside the frozen area.
			topLeftCell: formatAddress({
				row: Math.max(view.topLeft?.row ?? rows, rows),
				col: Math.max(view.topLeft?.col ?? cols, cols),
			}),
			activePane: pane,
			state: 'frozen',
		});
	}
	const split = view.split && (view.split.x > 0 || view.split.y > 0) ? view.split : undefined;
	if (split && !pane) {
		const { x, y } = split;
		pane = splitActivePane(split);
		inner += el('pane', {
			xSplit: x || undefined,
			ySplit: y || undefined,
			topLeftCell: view.topLeft ? formatAddress(view.topLeft) : undefined,
			activePane: pane,
			state: 'split',
		});
	}
	if (split) {
		// One selection per pane, in the order Excel writes them.
		for (const name of ['topLeft', 'topRight', 'bottomLeft', 'bottomRight'] as const) {
			const own = name === pane ? view.selection : split.paneSelections?.[name];
			if (own) inner += selectionXml(own, name);
		}
	} else if (view.selection) inner += selectionXml(view.selection, pane);
	const attributes = attrs({
		tabSelected: tabSelected || undefined,
		showGridLines: view.showGridLines ? undefined : false,
		showRowColHeaders: view.showHeaders ? undefined : false,
		showZeros: view.showZeros ? undefined : false,
		rightToLeft: view.rightToLeft || undefined,
		topLeftCell: split
			? split.topLeft
				? formatAddress(split.topLeft)
				: undefined
			: !view.freeze && view.topLeft
				? formatAddress(view.topLeft)
				: undefined,
		showFormulas: view.showFormulas || undefined,
		view: view.mode,
		zoomScale: view.zoom !== 100 ? view.zoom : undefined,
		zoomScaleNormal: view.zoom !== 100 ? view.zoom : undefined,
		workbookViewId: 0,
	});
	return `<sheetViews><sheetView${attributes}${inner ? `>${inner}</sheetView>` : '/>'}</sheetViews>`;
}

export function sheetFormatXml(sheet: Worksheet): string {
	const source = snapshotElement(sheet, 'sheetFormatPr');
	const current: ReturnType<typeof readSheetFormat> = { defaultRowHeight: sheet.defaultRowHeight };
	if (sheet.defaultColWidth !== undefined) current.defaultColWidth = sheet.defaultColWidth;
	if (sheet.format && Object.keys(sheet.format).length) current.format = sheet.format;
	// A loop, not a spread: a sheet can hold a row record per row, past a call's argument limit.
	let rowLevel = 0;
	for (const info of sheet.rowInfo.values()) rowLevel = Math.max(rowLevel, info.outlineLevel ?? 0);
	const colLevel = Math.max(0, ...sheet.columns.map((col) => col.outlineLevel ?? 0));
	if (
		source &&
		sameModel(readSheetFormat(source), current) &&
		Number(source.getAttribute('outlineLevelRow') ?? 0) === rowLevel &&
		Number(source.getAttribute('outlineLevelCol') ?? 0) === colLevel
	)
		return snapshotXml(sheet, 'sheetFormatPr') ?? '';
	const format = sheet.format ?? {};
	// Unknown and prefixed attributes (`x14ac:dyDescent`) are kept from the source.
	return `<sheetFormatPr${mergedAttrs(source, {
		baseColWidth: format.baseColWidth,
		defaultColWidth: sheet.defaultColWidth,
		defaultRowHeight: sheet.defaultRowHeight,
		customHeight: format.customHeight ?? (sheet.defaultRowHeight !== 15 ? true : undefined),
		zeroHeight: format.zeroHeight,
		thickTop: format.thickTop,
		thickBottom: format.thickBottom,
		outlineLevelRow: rowLevel || undefined,
		outlineLevelCol: colLevel || undefined,
	})}/>`;
}

export function colsXml(sheet: Worksheet, styleCount: number): string {
	const sorted = [...sheet.columns].sort((a, b) => a.min - b.min);
	let previous = -1;
	const out: string[] = [];
	for (const column of sorted) {
		const col: ColumnInfo = { ...column, min: Math.max(column.min, previous + 1) };
		if (col.max < col.min) continue;
		previous = col.max;
		const style = col.styleId && col.styleId < styleCount ? col.styleId : undefined;
		out.push(
			el('col', {
				min: col.min + 1,
				max: col.max + 1,
				width: col.width,
				style,
				hidden: col.hidden || undefined,
				bestFit: col.bestFit || undefined,
				customWidth: col.customWidth || undefined,
				outlineLevel: col.outlineLevel || undefined,
				collapsed: col.collapsed || undefined,
			}),
		);
	}
	return out.length ? `<cols>${out.join('')}</cols>` : '';
}

export function protectionXml(sheet: Worksheet): string {
	const protection = sheet.protection;
	if (!protection) return '';
	const source = snapshotElement(sheet, 'sheetProtection');
	if (source && sameModel(readProtection(source), protection))
		return snapshotXml(sheet, 'sheetProtection') ?? '';
	// Both hashes come from the model only: one the edit removed must not be copied back.
	const values: Record<string, string | boolean | undefined> = modernHashValues(
		protection.modernHash,
	);
	values['password'] = protection.passwordHash;
	values['sheet'] = protection.sheet || undefined;
	const allow = new Set(protection.allow ?? []);
	if (protection.sheet) {
		values['objects'] = !allow.has('objects');
		values['scenarios'] = !allow.has('scenarios');
	}
	for (const name of allow) values[name] = false;
	return el('sheetProtection', values);
}

const extLstOf = (source: XmlElement | undefined): string =>
	source
		? elements(source)
				.filter((node) => node.localName === 'extLst')
				.map((node) => outerXml(node))
				.join('')
		: '';

export function autoFilterXml(sheet: Worksheet, styles?: StyleWriter): string {
	const filter = sheet.autoFilter;
	if (!filter) return '';
	const source = snapshotElement(sheet, 'autoFilter');
	if (source && sameModel(unresolvedFilter(readAutoFilter(source)), unresolvedFilter(filter)))
		return snapshotXml(sheet, 'autoFilter') ?? '';
	const width = filter.range.end.col - filter.range.start.col + 1;
	const columns = (filter.columns ?? [])
		.filter((column) => column.offset >= 0 && column.offset < width)
		.sort((a, b) => a.offset - b.offset)
		.map((column) => filterColumnXml(column, styles))
		.join('');
	// A kept `sortState` (inside the filter) stays only while the range is unchanged.
	const sort = filter.sortState
		? sortStateXml(filter.sortState, styles) + extLstOf(source)
		: source && sameModel(readAutoFilter(source)?.range, filter.range)
			? elements(source)
					.filter((node) => node.localName === 'sortState' || node.localName === 'extLst')
					.map((node) => outerXml(node))
					.join('')
			: '';
	return el('autoFilter', { ref: formatRange(filter.range) }, columns + sort);
}

export function mergeCellsXml(sheet: Worksheet): string {
	if (!sheet.merges.length) return '';
	const cells = sheet.merges.map((range) => `<mergeCell ref="${formatRange(range)}"/>`).join('');
	return `<mergeCells count="${sheet.merges.length}">${cells}</mergeCells>`;
}
