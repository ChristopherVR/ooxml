import type { ChartSpace } from '../../chart/model';
import { NS, buildXml, children, first, parseXml } from '../../xml/index';
import type { ChartObject, PivotTable, Workbook } from '../model';

const C14 = 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart';
const PIVOT_OPTIONS = '{781A3756-C4B2-4CAC-9D66-4F8BD8637D16}';
const buttons = {
	filter: 'dropZoneFilter',
	value: 'dropZoneData',
	legend: 'dropZoneSeries',
	axis: 'dropZoneCategories',
} as const;
const visibility = new WeakMap<ChartSpace, ReadonlySet<PivotChartButton>>();

export type PivotChartButton = keyof typeof buttons;

export function pivotChartSource(space: Pick<ChartSpace, 'pivotSource'>):
	| {
			prefix: string;
			sheet: string;
			pivot: string;
	  }
	| undefined {
	const source = space.pivotSource?.name;
	if (!source?.startsWith('[')) return undefined;
	const bookEnd = source.indexOf(']');
	const separator = source.lastIndexOf('!');
	if (bookEnd < 0 || separator <= bookEnd + 1 || separator === source.length - 1) return undefined;
	return {
		prefix: source.slice(0, bookEnd + 1),
		sheet: source.slice(bookEnd + 1, separator),
		pivot: source.slice(separator + 1),
	};
}

export function chartUsesPivot(chart: ChartObject, sheetName: string, pivotName: string): boolean {
	const source = chart.chartSpace && pivotChartSource(chart.chartSpace);
	return source?.sheet === sheetName && source.pivot === pivotName;
}

export function pivotChartTable(
	workbook: Workbook,
	chart: ChartObject,
):
	| {
			sheet: number;
			pivot: PivotTable;
	  }
	| undefined {
	const source = chart.chartSpace && pivotChartSource(chart.chartSpace);
	if (!source) return undefined;
	const sheet = workbook.sheets.findIndex((entry) => entry.name === source.sheet);
	const pivot = workbook.sheets[sheet]?.pivotTables?.find((entry) => entry.name === source.pivot);
	return pivot ? { sheet, pivot } : undefined;
}

function pivotOptions(xml: string | undefined) {
	const doc = parseXml(xml ?? `<c:extLst xmlns:c="${NS.c}"/>`);
	const root = doc.documentElement;
	const ext = children(root, 'ext', NS.c).find(
		(entry) => entry.getAttribute('uri') === PIVOT_OPTIONS,
	);
	return { doc, root, ext, options: first(ext, 'pivotOptions', C14) };
}

export function pivotChartButtons(chart: ChartObject): ReadonlySet<PivotChartButton> {
	const space = chart.chartSpace;
	if (!space?.pivotSource) return new Set();
	const cached = visibility.get(space);
	if (cached) return cached;
	const { options } = pivotOptions(space.extLst);
	const on = (name: string) =>
		['1', 'true'].includes(first(options, name, C14)?.getAttribute('val') ?? '');
	const roles: PivotChartButton[] = ['filter', 'axis', 'legend', 'value'];
	const result = new Set(
		roles.filter((role) => !options || (on('dropZonesVisible') && on(buttons[role]))),
	);
	visibility.set(space, result);
	return result;
}

export function withoutPivotChartButtons(space: ChartSpace): void {
	if (!space.extLst) return;
	const { root, ext } = pivotOptions(space.extLst);
	if (!ext) return;
	root.removeChild(ext);
	if (children(root, 'ext', NS.c).length) space.extLst = buildXml(root);
	else delete space.extLst;
}

export function chartSpaceWithPivotButtons(space: ChartSpace, visible: boolean): ChartSpace {
	const { doc, root, ext, options } = pivotOptions(space.extLst);
	const settings = doc.createElementNS(C14, 'c14:pivotOptions');
	const add = (name: string, value: boolean) => {
		const element = doc.createElementNS(C14, `c14:${name}`);
		element.setAttribute('val', value ? '1' : '0');
		settings.appendChild(element);
	};
	if (visible) for (const name of Object.values(buttons)) add(name, true);
	add('dropZonesVisible', visible);
	const extension = ext ?? doc.createElementNS(NS.c, 'c:ext');
	if (options) extension.removeChild(options);
	extension.appendChild(settings);
	if (!ext) {
		extension.setAttribute('uri', PIVOT_OPTIONS);
		root.appendChild(extension);
	}
	return { ...space, extLst: buildXml(root) };
}
