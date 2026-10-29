// Reads `customWorkbookViews` and a sheet's `customSheetViews` (see `model-custom-views.ts`).
// The sheet-level view reuses the readers of the sheet itself: the attributes and panes of a
// `customSheetView` are the ones of a `sheetView`, and its page setup children are the sheet's.
import { NS, elements, parseXml, type XmlElement } from '../../xml/index';
import type { CustomSheetView, CustomWorkbookView } from '../model-custom-views';
import { readPrintOptions } from './print-options';
import { readAutoFilter, readPageSetup, readSheetView } from './sheet-props';
import { att, boolAttr, outerXml, xChildren } from './xml-util';

/** `customSheetView` attributes the model reads; the rest are kept in `attrs`. */
const SHEET_VIEW_ATTRS = [
	'guid',
	'scale',
	'showGridLines',
	'showRowCol',
	'zeroValues',
	'topLeftCell',
	'view',
	'showFormulas',
	'hiddenRows',
	'hiddenColumns',
	'filter',
	'showAutoFilter',
	'fitToPage',
] as const;

const WORKBOOK_VIEW_KNOWN = ['name', 'guid', 'includePrintSettings', 'includeHiddenRowCol'];

export function readCustomWorkbookViews(node: XmlElement): CustomWorkbookView[] {
	return xChildren(node, 'customWorkbookView').flatMap((view) => {
		const guid = att(view, 'guid');
		if (!guid) return [];
		const attrs: Record<string, string> = {};
		for (const attribute of Array.from(view.attributes))
			if (!WORKBOOK_VIEW_KNOWN.includes(attribute.name) && !attribute.name.startsWith('xmlns'))
				attrs[attribute.name] = attribute.value;
		return [
			{
				guid,
				name: att(view, 'name') ?? '',
				includePrintSettings: boolAttr(view, 'includePrintSettings', true),
				includeHiddenRowCol: boolAttr(view, 'includeHiddenRowCol', true),
				attrs,
			},
		];
	});
}

const escape = (value: string) => value.replace(/&/g, '&amp;').replace(/"/g, '&quot;');

/** The view as a `sheetViews` element, so `readSheetView` interprets it. */
function asSheetViews(view: XmlElement): XmlElement {
	const map: Array<[from: string, to: string]> = [
		['scale', 'zoomScale'],
		['showGridLines', 'showGridLines'],
		['showRowCol', 'showRowColHeaders'],
		['zeroValues', 'showZeros'],
		['topLeftCell', 'topLeftCell'],
		['view', 'view'],
		['showFormulas', 'showFormulas'],
	];
	const attributes = map
		.flatMap(([from, to]) => {
			const value = att(view, from);
			return value === undefined ? [] : [` ${to}="${escape(value)}"`];
		})
		.join('');
	const children = elements(view)
		.filter((child) => child.localName === 'pane' || child.localName === 'selection')
		.map((child) => outerXml(child))
		.join('');
	return parseXml(
		`<sheetViews xmlns="${NS.x}"><sheetView${attributes}>${children}</sheetView></sheetViews>`,
	).documentElement;
}

export function readCustomSheetViews(node: XmlElement): CustomSheetView[] {
	return xChildren(node, 'customSheetView').flatMap((view) => {
		const guid = att(view, 'guid');
		if (!guid) return [];
		const child = (name: string) => xChildren(view, name)[0];
		const pageSetup = readPageSetup(
			child('pageMargins'),
			child('pageSetup'),
			child('headerFooter'),
			boolAttr(view, 'fitToPage', false),
		);
		const printOptions = readPrintOptions(child('printOptions'));
		const filter = readAutoFilter(child('autoFilter'));
		const attrs: Record<string, string> = {};
		for (const attribute of Array.from(view.attributes))
			if (
				!(SHEET_VIEW_ATTRS as readonly string[]).includes(attribute.name) &&
				!attribute.name.startsWith('xmlns')
			)
				attrs[attribute.name] = attribute.value;
		const modelled = [
			'pane',
			'selection',
			'pageMargins',
			'printOptions',
			'pageSetup',
			'headerFooter',
			'autoFilter',
		];
		return [
			{
				guid,
				view: readSheetView(asSheetViews(view)),
				...(boolAttr(view, 'hiddenRows', false) ? { hiddenRows: true as const } : {}),
				...(boolAttr(view, 'hiddenColumns', false) ? { hiddenColumns: true as const } : {}),
				...(filter ? { filter } : {}),
				...(boolAttr(view, 'showAutoFilter', false) ? { showAutoFilter: true as const } : {}),
				...(pageSetup ? { pageSetup } : {}),
				...(printOptions ? { printOptions } : {}),
				attrs,
				extra: elements(view)
					.filter((c) => !modelled.includes(c.localName ?? ''))
					.map((c) => outerXml(c)),
			},
		];
	});
}
