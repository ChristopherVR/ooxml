/**
 * `c:chartSpace`-level formatting: the chart area `c:spPr` and
 * `c:roundedCorners`, positioned by the `CT_ChartSpace` sequence (`date1904`,
 * `lang`, `roundedCorners`, `AlternateContent`, `style`, `clrMapOvr`,
 * `pivotSource`, `protection`, `chart`, `spPr`, `txPr`, `externalData`,
 * `printSettings`, `userShapes`, `extLst`), and the plot area `c:spPr`
 * (`CT_PlotArea`: ... `dTable`, `spPr`, `extLst`).
 *
 * @module utils/chart-space-format
 */
import type { PptxChartData, XmlObject } from '../types';
import type { ChartAreaFormat, ChartAreaFormatWriteOptions } from './chart-area-format';
import { applyChartAreaFormatToXml, buildChartAreaSpPr } from './chart-area-format';

type GetLocalName = (key: string) => string;
type XmlValue = XmlObject[string];

/** `CT_ChartSpace` children that follow `c:roundedCorners`. */
const AFTER_ROUNDED_CORNERS = new Set([
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
]);
/** `CT_ChartSpace` children that follow `c:spPr`. */
const AFTER_CHART_SPACE_SP_PR = new Set([
	'txPr',
	'externalData',
	'printSettings',
	'userShapes',
	'extLst',
]);

function findKey(obj: XmlObject, local: string, getLocalName: GetLocalName): string | undefined {
	return Object.keys(obj).find((key) => getLocalName(key) === local);
}

/** Insert `key: value` before the first child whose local name is in `after`. */
function insertBefore(
	node: XmlObject,
	key: string,
	value: XmlValue,
	after: Set<string>,
	getLocalName: GetLocalName,
): void {
	const entries = Object.entries(node);
	const at = entries.findIndex(
		([candidate]) => !candidate.startsWith('@_') && after.has(getLocalName(candidate)),
	);
	entries.splice(at === -1 ? entries.length : at, 0, [key, value]);
	for (const candidate of Object.keys(node)) {
		delete node[candidate];
	}
	for (const [candidate, child] of entries) {
		node[candidate] = child;
	}
}

/** The chart-area format from the model's style. */
export function chartAreaFormatOf(chartData: PptxChartData): ChartAreaFormat {
	const style = chartData.style;
	return {
		...(style?.chartAreaFill !== undefined ? { fill: style.chartAreaFill } : {}),
		...(style?.chartAreaGradient ? { gradient: style.chartAreaGradient } : {}),
		...(style?.chartAreaBorder !== undefined ? { border: style.chartAreaBorder } : {}),
	};
}

/** The plot-area format from the model's style. */
export function plotAreaFormatOf(chartData: PptxChartData): ChartAreaFormat {
	const style = chartData.style;
	return {
		...(style?.plotAreaFill !== undefined ? { fill: style.plotAreaFill } : {}),
		...(style?.plotAreaGradient ? { gradient: style.plotAreaGradient } : {}),
		...(style?.plotAreaBorder !== undefined ? { border: style.plotAreaBorder } : {}),
	};
}

/**
 * Add the chart-area `c:spPr`, the plot-area `c:spPr` and `c:roundedCorners`
 * to a generated chart part. Each is written only when the model sets it, so
 * a chart created without them is generated exactly as before.
 */
export function applyGeneratedChartSpaceFormat(tree: XmlObject, chartData: PptxChartData): void {
	const chartSpace = tree['c:chartSpace'] as XmlObject | undefined;
	if (!chartSpace) {
		return;
	}
	const localName = (key: string) => key.replace(/^.*:/u, '');
	const plotArea = (chartSpace['c:chart'] as XmlObject | undefined)?.['c:plotArea'] as
		| XmlObject
		| undefined;
	const plotSpPr = buildChartAreaSpPr(plotAreaFormatOf(chartData));
	if (plotArea && plotSpPr) {
		insertBefore(plotArea, 'c:spPr', plotSpPr, new Set(['extLst']), localName);
	}
	const chartSpPr = buildChartAreaSpPr(chartAreaFormatOf(chartData));
	if (chartSpPr) {
		insertBefore(chartSpace, 'c:spPr', chartSpPr, AFTER_CHART_SPACE_SP_PR, localName);
	}
	if (chartData.roundedCorners !== undefined) {
		insertBefore(
			chartSpace,
			'c:roundedCorners',
			{ '@_val': chartData.roundedCorners ? '1' : '0' },
			AFTER_ROUNDED_CORNERS,
			localName,
		);
	}
}

function authoredBool(node: unknown): boolean | undefined {
	if (node === undefined) {
		return undefined;
	}
	const val = node && typeof node === 'object' ? (node as XmlObject)['@_val'] : undefined;
	if (val === undefined || val === null || val === '') {
		return true;
	}
	return !(val === '0' || val === 'false');
}

/**
 * Reconcile a loaded chart's `c:roundedCorners`, chart-area `c:spPr` and
 * plot-area `c:spPr` with the model. Untouched parts stay as authored; a
 * cleared `roundedCorners` removes the element. Note that PowerPoint draws
 * rounded corners when the element is absent, so square corners need an
 * explicit `false`.
 *
 * @returns Whether anything changed.
 */
export function applyChartSpaceFormatToXml(
	chartSpace: XmlObject,
	plotArea: XmlObject | undefined,
	chartData: PptxChartData,
	getLocalName: GetLocalName,
	options: Omit<ChartAreaFormatWriteOptions, 'insertSpPr'>,
): boolean {
	let changed = false;
	const roundedKey = findKey(chartSpace, 'roundedCorners', getLocalName);
	const authoredRounded = roundedKey ? authoredBool(chartSpace[roundedKey]) : undefined;
	if (chartData.roundedCorners === undefined) {
		if (roundedKey) {
			delete chartSpace[roundedKey];
			changed = true;
		}
	} else if (chartData.roundedCorners !== authoredRounded) {
		const node = { '@_val': chartData.roundedCorners ? '1' : '0' };
		if (roundedKey) {
			chartSpace[roundedKey] = node;
		} else {
			insertBefore(chartSpace, 'c:roundedCorners', node, AFTER_ROUNDED_CORNERS, getLocalName);
		}
		changed = true;
	}
	changed =
		applyChartAreaFormatToXml(chartSpace, chartAreaFormatOf(chartData), getLocalName, {
			...options,
			insertSpPr: (owner, spPr) =>
				insertBefore(owner, 'c:spPr', spPr, AFTER_CHART_SPACE_SP_PR, getLocalName),
		}) || changed;
	if (plotArea) {
		changed =
			applyChartAreaFormatToXml(plotArea, plotAreaFormatOf(chartData), getLocalName, {
				...options,
				insertSpPr: (owner, spPr) =>
					insertBefore(owner, 'c:spPr', spPr, new Set(['extLst']), getLocalName),
			}) || changed;
	}
	return changed;
}
