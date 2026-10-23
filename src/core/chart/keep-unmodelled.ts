// Saving an edited chart part over its source: `writeChartSpace` writes what `ChartSpace` models,
// and the direct children of `c:chartSpace` and `c:chart` the parser reports as not modelled
// (`c:clrMapOvr`, `c:pivotSource`, `c:protection`, `c:pivotFmts`) are carried over from the source
// part at their schema positions, so a pivot chart or a protected chart keeps them through an edit.
// Unmodelled content deeper in the part (inside the plot area, a series or an axis) cannot be
// carried this way: `keep-source.ts` says when a part has some, and such a part keeps its bytes.
import { NS, buildXml, elements, first, parseXml, type XmlElement } from '../xml/index';

/** CT_ChartSpace children in schema order. */
const SPACE_ORDER = [
	'date1904',
	'lang',
	'roundedCorners',
	'AlternateContent',
	'style',
	'clrMapOvr',
	'pivotSource',
	'protection',
	'chart',
	'spPr',
	'txPr',
	'externalData',
	'printSettings',
	'userShapes',
	'extLst',
];

/** CT_Chart children in schema order. */
const CHART_ORDER = [
	'title',
	'autoTitleDeleted',
	'pivotFmts',
	'view3D',
	'floor',
	'sideWall',
	'backWall',
	'plotArea',
	'legend',
	'plotVisOnly',
	'dispBlanksAs',
	'showDLblsOverMax',
	'extLst',
];

/** The children `ChartSpace` does not model, by parent. */
const UNMODELLED: Record<'space' | 'chart', ReadonlySet<string>> = {
	space: new Set(['clrMapOvr', 'pivotSource', 'protection']),
	chart: new Set(['pivotFmts']),
};

/** Copies the unmodelled children of `from` into `to` (same element in the written part). */
function graft(from: XmlElement, to: XmlElement, order: string[], kinds: ReadonlySet<string>) {
	let changed = false;
	for (const child of elements(from)) {
		if (child.namespaceURI !== NS.c || !kinds.has(child.localName)) continue;
		if (first(to, child.localName, NS.c)) continue;
		const rank = order.indexOf(child.localName);
		const before = elements(to).find((item) => order.indexOf(item.localName) > rank);
		to.insertBefore(to.ownerDocument.importNode(child, true), before ?? null);
		changed = true;
	}
	return changed;
}

/**
 * `written` (a part `writeChartSpace` wrote for an edited model of `source`) with the source's
 * unmodelled `c:chartSpace` and `c:chart` children restored; `written` itself when there are none.
 */
export function withUnmodelledChartChildren(source: string, written: string): string {
	if (!/<(?:\w+:)?(?:clrMapOvr|pivotSource|protection|pivotFmts)\b/.test(source)) return written;
	const from = parseXml(source, { label: 'Chart part' }).documentElement;
	const doc = parseXml(written, { label: 'Chart part' });
	const to = doc.documentElement;
	let changed = graft(from, to, SPACE_ORDER, UNMODELLED.space);
	const fromChart = first(from, 'chart', NS.c);
	const toChart = first(to, 'chart', NS.c);
	if (fromChart && toChart)
		changed = graft(fromChart, toChart, CHART_ORDER, UNMODELLED.chart) || changed;
	if (!changed) return written;
	const declaration = /^<\?xml[^>]*\?>(\r\n|\n)?/.exec(written)?.[0] ?? '';
	return declaration + buildXml(to);
}
