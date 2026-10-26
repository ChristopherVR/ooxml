// The Office chart15 value-from-cells extension, edited without replacing sibling extensions.
import { NS, buildXml, children, elements, parseXml, type XmlElement } from '../../xml/index';
import type { ChartSpace } from '../model';
import type { ChartDataLabels } from '../model-series';
import { CHART_ROOT_BINDINGS, stripDeclarations } from '../xml-fragment';
import { chartSeriesEntryAt, editChartSpace } from './draft';

const C15 = 'http://schemas.microsoft.com/office/drawing/2012/chart';
const URI = '{CE6537A1-D6FC-4f65-9D91-7224C49458BB}';

export interface ChartDataLabelsRange {
	formula: string;
	cache: readonly string[];
}

function extensionRoot(xml: string | undefined): XmlElement {
	return parseXml(
		`<c:wrapper xmlns:c="${NS.c}" xmlns:c15="${C15}">${xml ?? '<c:extLst/>'}</c:wrapper>`,
	).documentElement.getElementsByTagNameNS(NS.c, 'extLst')[0]!;
}

/** Values and visibility as authored, independent of any host's workbook. */
export function chartDataLabelsRange(
	labels: ChartDataLabels,
): (ChartDataLabelsRange & { show: boolean }) | undefined {
	if (!labels.extLst) return undefined;
	const root = extensionRoot(labels.extLst);
	const ext = children(root, 'ext', NS.c).find((node) => node.getAttribute('uri') === URI);
	if (!ext) return undefined;
	const range = children(ext, 'datalabelsRange', C15)[0];
	const formula = range && children(range, 'f', C15)[0]?.textContent;
	if (!formula) return undefined;
	const cache = children(range!, 'dlblRangeCache', C15)[0];
	const count = Number(cache && children(cache, 'ptCount', NS.c)[0]?.getAttribute('val'));
	const points = cache ? children(cache, 'pt', NS.c) : [];
	const indices = points.map((pt) => Number(pt.getAttribute('idx')));
	const length = indices.reduce(
		(max, index) => (Number.isSafeInteger(index) && index >= 0 ? Math.max(max, index + 1) : max),
		Number.isSafeInteger(count) && count >= 0 ? count : 0,
	);
	if (length > 1_000_000)
		throw new RangeError('Chart label cache exceeds the supported point limit');
	const values = Array.from({ length }, () => '');
	points.forEach((pt, i) => {
		const index = indices[i]!;
		if (Number.isSafeInteger(index) && index >= 0)
			values[index] = children(pt, 'v', NS.c)[0]?.textContent ?? '';
	});
	const show = children(ext, 'showDataLabelsRange', C15)[0]?.getAttribute('val');
	return { formula, cache: values, show: show === '1' || show === 'true' };
}

function patchRange(
	labels: ChartDataLabels,
	range: ChartDataLabelsRange | null,
	show: boolean,
): void {
	const previous = chartDataLabelsRange(labels);
	if (
		(!range && !previous) ||
		(range &&
			previous &&
			range.formula === previous.formula &&
			show === previous.show &&
			JSON.stringify(range.cache) === JSON.stringify(previous.cache))
	)
		return;
	const root = extensionRoot(labels.extLst);
	let ext = children(root, 'ext', NS.c).find((node) => node.getAttribute('uri') === URI);
	if (!ext) {
		ext = root.ownerDocument.createElementNS(NS.c, 'c:ext');
		ext.setAttribute('uri', URI);
		root.appendChild(ext);
	}
	for (const child of elements(ext))
		if (
			child.namespaceURI === C15 &&
			['datalabelsRange', 'showDataLabelsRange'].includes(child.localName)
		)
			ext.removeChild(child);
	const add = (parent: XmlElement, ns: string, name: string, text?: string) => {
		const node = root.ownerDocument.createElementNS(ns, name);
		if (text !== undefined) node.textContent = text;
		parent.appendChild(node);
		return node;
	};
	if (range) {
		add(ext, C15, 'c15:showDataLabelsRange').setAttribute('val', show ? '1' : '0');
		const node = add(ext, C15, 'c15:datalabelsRange');
		add(node, C15, 'c15:f', range.formula);
		const cache = add(node, C15, 'c15:dlblRangeCache');
		add(cache, NS.c, 'c:ptCount').setAttribute('val', String(range.cache.length));
		range.cache.forEach((text, index) => {
			const pt = add(cache, NS.c, 'c:pt');
			pt.setAttribute('idx', String(index));
			add(pt, NS.c, 'c:v', text);
		});
	}
	if (!elements(ext).length) root.removeChild(ext);
	if (elements(root).length) labels.extLst = stripDeclarations(buildXml(root), CHART_ROOT_BINDINGS);
	else delete labels.extLst;
}

/** Changes cached display labels only; a host owns formula evaluation and workbook relationships. */
export function setChartDataLabelsRange(
	space: ChartSpace,
	range: ChartDataLabelsRange | null,
	show = true,
	target: { series?: number } = {},
): ChartSpace {
	if (range && (!range.formula.trim() || range.cache.length > 1_000_000))
		throw new RangeError('Chart label range needs a formula and a bounded cache');
	return editChartSpace(space, (draft) => {
		const holders =
			target.series === undefined
				? draft.plotArea.groups.filter((group) => group.kind !== 'surface')
				: [chartSeriesEntryAt(draft, target.series).series];
		for (const holder of holders) {
			if (!holder.dataLabels && !range) continue;
			patchRange((holder.dataLabels ??= { labels: [] }), range, show);
		}
	});
}
