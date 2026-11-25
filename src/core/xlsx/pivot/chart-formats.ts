import { NS, elements, first, parseXml, type XmlElement } from '../../xml/index';
import type { PivotTable } from '../model';
import type { PivotGrid } from './aggregate';
import { pivotChartColumns } from './chart-series';

export function updatePivotChartFormats(
	root: XmlElement,
	pivot: PivotTable,
	grid: PivotGrid,
	columnFields: number[],
	columnItems: number[][],
	charts: ReadonlyMap<number, number>,
): void {
	if (!charts.size) return;
	const columns = pivotChartColumns(pivot, grid);
	const existing = first(root, 'chartFormats', NS.x);
	const kept = existing
		? elements(existing).filter((entry) => !charts.has(Number(entry.getAttribute('chart'))))
		: [];
	let xml = '';
	let count = kept.length;
	for (const [chart, size] of charts) {
		columns.slice(0, size).forEach(({ entry, measure }, format) => {
			const refs = [
				[4294967294, measure],
				...entry.path.map((value, level) => [columnFields[level], columnItems[level]?.[value]]),
			];
			xml += `<chartFormat chart="${chart}" format="${format}" series="1"><pivotArea type="data" outline="0" fieldPosition="0"><references count="${refs.length}">${refs.map(([field, item]) => `<reference field="${field}" count="1" selected="0"><x v="${item}"/></reference>`).join('')}</references></pivotArea></chartFormat>`;
			count++;
		});
	}
	const formats = parseXml(
		`<chartFormats xmlns="${NS.x}" count="${count}">${xml}</chartFormats>`,
	).documentElement;
	const doc = root.ownerDocument;
	const next = doc.importNode(formats, true);
	for (const entry of kept) next.appendChild(doc.importNode(entry, true));
	if (existing) root.removeChild(existing);
	const before = elements(root).find((node) =>
		['pivotTableStyleInfo', 'filters', 'extLst'].includes(node.localName),
	);
	if (count) root.insertBefore(next, before ?? null);
	root.setAttribute(
		'chartFormat',
		String(Math.max(Number(root.getAttribute('chartFormat')) || 0, 1 + Math.max(...charts.keys()))),
	);
}
