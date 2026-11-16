// Series hidden by a chart filter ([MS-ODRAWXML] 2.3.4: `c15:filteredBarSeries`,
// `c15:filteredLineSeries`... in a chart group's extension list, under the extension URI
// {02D57815-91ED-43cb-92C2-25804820EDAC}). A hidden series keeps its data, references and
// formatting inside the extension, as PowerPoint and Excel write it when the Chart Filters pane
// clears a series: hiding moves the series there, restoring moves it back to its group, in `c:idx`
// order, and removing drops it.
import { NS, buildXml, elements, first, parseXml, type XmlElement } from '../../xml/index';
import type { ChartGroupKind, ChartPlotGroup, ChartSpace } from '../model';
import type { ChartSpaceSeries } from '../model-series';
import { parseSeries } from '../parse-series';
import { seriesXml } from '../write-series';
import { CHART_ROOT_BINDINGS } from '../xml-fragment';
import { chartSeriesEntries, editChartSpace } from './draft';

const C15 = 'http://schemas.microsoft.com/office/drawing/2012/chart';
const FILTERED_URI = '{02D57815-91ED-43cb-92C2-25804820EDAC}';

/** The filtered series element of each group family ([MS-ODRAWXML] CT_FilteredXxxSeries). */
const WRAPPER: Readonly<Record<ChartGroupKind, string>> = {
	bar: 'filteredBarSeries',
	line: 'filteredLineSeries',
	stock: 'filteredLineSeries',
	area: 'filteredAreaSeries',
	pie: 'filteredPieSeries',
	doughnut: 'filteredPieSeries',
	ofPie: 'filteredPieSeries',
	scatter: 'filteredScatterSeries',
	radar: 'filteredRadarSeries',
	bubble: 'filteredBubbleSeries',
	surface: 'filteredSurfaceSeries',
};

const EMPTY_EXTENSIONS = `<c:extLst xmlns:c="${NS.c}"/>`;

function extensionList(group: ChartPlotGroup): XmlElement {
	return parseXml(group.extLst ?? EMPTY_EXTENSIONS, { label: 'Chart extension list' })
		.documentElement;
}

/** The hidden series elements (`c15:ser`) of `extLst`, with their wrapper. */
function filteredEntries(extLst: XmlElement): XmlElement[] {
	return elements(extLst)
		.filter((ext) => ext.localName === 'ext' && ext.getAttribute('uri') === FILTERED_URI)
		.flatMap((ext) =>
			elements(ext).filter((wrapper) => /^filtered\w*Series$/u.test(wrapper.localName)),
		)
		.flatMap((wrapper) => elements(wrapper).filter((item) => item.localName === 'ser'));
}

const entryIndex = (entry: XmlElement): number | undefined => {
	const value = first(entry, 'idx', NS.c)?.getAttribute('val');
	return value === null || value === undefined ? undefined : Number(value);
};

/** Writes `extLst` back to the group, dropping wrappers, the extension and the list once empty. */
function storeExtensions(group: ChartPlotGroup, extLst: XmlElement): void {
	for (const ext of elements(extLst)) {
		if (ext.getAttribute('uri') !== FILTERED_URI) continue;
		for (const wrapper of elements(ext))
			if (/^filtered\w*Series$/u.test(wrapper.localName) && elements(wrapper).length === 0)
				ext.removeChild(wrapper);
		if (elements(ext).length === 0) extLst.removeChild(ext);
	}
	if (elements(extLst).length === 0) delete group.extLst;
	else group.extLst = buildXml(extLst);
}

const REFERENCES = new Set(['strRef', 'numRef', 'multiLvlStrRef']);

function references(entry: XmlElement): XmlElement[] {
	const found: XmlElement[] = [];
	const visit = (element: XmlElement) => {
		for (const child of elements(element)) {
			if (child.namespaceURI === NS.c && REFERENCES.has(child.localName)) found.push(child);
			else if (child.localName !== 'extLst') visit(child);
		}
	};
	visit(entry);
	return found;
}

/**
 * A hidden series names its cells in `c15:formulaRef/c15:sqref` instead of `c:f`, as Excel 16 writes
 * a series its Chart Filters (`Series.IsFiltered`) hide, so the cells are not recalculated into it.
 */
function formulasToReferences(entry: XmlElement): void {
	const doc = entry.ownerDocument;
	for (const reference of references(entry)) {
		const formula = first(reference, 'f', NS.c);
		if (!formula) continue;
		const extLst = doc.createElementNS(NS.c, 'c:extLst');
		const ext = doc.createElementNS(NS.c, 'c:ext');
		ext.setAttribute('uri', FILTERED_URI);
		const ref = doc.createElementNS(C15, 'c15:formulaRef');
		const sqref = doc.createElementNS(C15, 'c15:sqref');
		sqref.textContent = formula.textContent ?? '';
		ref.appendChild(sqref);
		ext.appendChild(ref);
		extLst.appendChild(ext);
		reference.replaceChild(extLst, formula);
	}
}

/** The reverse of {@link formulasToReferences}, for a series shown again. */
function referencesToFormulas(entry: XmlElement): void {
	const doc = entry.ownerDocument;
	for (const reference of references(entry)) {
		const extLst = first(reference, 'extLst', NS.c);
		if (!extLst || first(reference, 'f', NS.c)) continue;
		const ext = elements(extLst).find((item) => item.getAttribute('uri') === FILTERED_URI);
		const sqref = ext && first(first(ext, 'formulaRef', C15), 'sqref', C15);
		if (!ext || !sqref) continue;
		const formula = doc.createElementNS(NS.c, 'c:f');
		formula.textContent = sqref.textContent ?? '';
		reference.insertBefore(formula, reference.firstChild);
		extLst.removeChild(ext);
		if (elements(extLst).length === 0) reference.removeChild(extLst);
	}
}

/** The `c:idx` of every series a filter hides, in document order. */
export function chartFilteredSeriesIndexes(space: ChartSpace): number[] {
	return space.plotArea.groups.flatMap((group) =>
		group.extLst
			? filteredEntries(extensionList(group)).flatMap((entry) => {
					const index = entryIndex(entry);
					return index === undefined ? [] : [index];
				})
			: [],
	);
}

/** Hides the series whose `c:idx` is `index`: it moves into its group's filtered series. */
export function filterChartSeries(space: ChartSpace, index: number): ChartSpace {
	return editChartSpace(space, (draft) => {
		const entry = chartSeriesEntries(draft).find(({ series }) => series.index === index);
		if (!entry) throw new RangeError(`No chart series with idx ${index}`);
		const { group, position } = entry;
		// PowerPoint's Chart Filters keep one plotted series in each group.
		if (group.series.length < 2)
			throw new RangeError('Hiding the last series of a chart group is not supported');
		const extLst = extensionList(group);
		const doc = extLst.ownerDocument;
		let ext = elements(extLst).find(
			(item) => item.localName === 'ext' && item.getAttribute('uri') === FILTERED_URI,
		);
		if (!ext) {
			ext = doc.createElementNS(NS.c, 'c:ext');
			ext.setAttribute('uri', FILTERED_URI);
			ext.setAttributeNS('http://www.w3.org/2000/xmlns/', 'xmlns:c15', C15);
			extLst.appendChild(ext);
		}
		const local = WRAPPER[group.kind];
		let wrapper = elements(ext).find((item) => item.localName === local);
		if (!wrapper) {
			wrapper = doc.createElementNS(C15, `c15:${local}`);
			ext.appendChild(wrapper);
		}
		const declarations = [...CHART_ROOT_BINDINGS]
			.map(([prefix, uri]) => ` xmlns:${prefix}="${uri}"`)
			.join('');
		const written = seriesXml({ bindings: CHART_ROOT_BINDINGS }, entry.series, position)
			.replace(/^<c:ser>/u, `<c15:ser xmlns:c15="${C15}"${declarations}>`)
			.replace(/<\/c:ser>$/u, '</c15:ser>');
		const hidden = parseXml(written, { label: 'Chart series' }).documentElement;
		formulasToReferences(hidden);
		const moved = doc.importNode(hidden, true) as XmlElement;
		// The extension declares `c15`; the series repeats none of the declarations it was parsed with.
		for (const attribute of [...moved.attributes])
			if (attribute.name.startsWith('xmlns')) moved.removeAttribute(attribute.name);
		wrapper.appendChild(moved);
		group.series.splice(position, 1);
		storeExtensions(group, extLst);
	});
}

/** Takes the hidden series `index` out of every group's filtered series; returns the group. */
function takeEntry(draft: ChartSpace, index: number) {
	for (const group of draft.plotArea.groups) {
		if (!group.extLst) continue;
		const extLst = extensionList(group);
		const entry = filteredEntries(extLst).find((item) => entryIndex(item) === index);
		if (!entry) continue;
		entry.parentNode!.removeChild(entry);
		storeExtensions(group, extLst);
		return { group, entry };
	}
	throw new RangeError(`No hidden chart series with idx ${index}`);
}

/** Shows the hidden series `index` again, in its group at its `c:idx` position. */
export function restoreChartFilteredSeries(space: ChartSpace, index: number): ChartSpace {
	return editChartSpace(space, (draft) => {
		const { group, entry } = takeEntry(draft, index);
		referencesToFormulas(entry);
		const series: ChartSpaceSeries = parseSeries({ issues: [] }, entry);
		const at = group.series.findIndex((item) => (item.index ?? 0) > index);
		group.series.splice(at < 0 ? group.series.length : at, 0, series);
	});
}

/** Drops the hidden series `index` (its data leaves the part). */
export function removeChartFilteredSeries(space: ChartSpace, index: number): ChartSpace {
	return editChartSpace(space, (draft) => {
		takeEntry(draft, index);
	});
}
