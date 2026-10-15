/**
 * Target keys for chart element formats and the save-time writer that finds
 * each target's owner node in a chart part and applies the recorded format
 * to its `c:spPr`.
 *
 * @module utils/chart-element-format-targets
 */
import type { PptxChartData, PptxChartElementFormat, XmlObject } from '../types';
import { parseChartFormatTargetKey } from './chart-element-format-keys';
import { applyChartElementFormatToSpPr } from './chart-element-format-xml';

export { chartFormatTargetKey, parseChartFormatTargetKey } from './chart-element-format-keys';

type GetLocalName = (key: string) => string;

const CHART_SPACE_AFTER_SP_PR = ['txPr', 'externalData', 'printSettings', 'userShapes', 'extLst'];
const TEXT_OWNER_AFTER_SP_PR = ['txPr', 'extLst'];
const AXIS_AFTER_SP_PR = [
	'txPr',
	'crossAx',
	'crosses',
	'crossesAt',
	'crossBetween',
	'auto',
	'lblAlgn',
	'lblOffset',
	'majorUnit',
	'minorUnit',
	'dispUnits',
	'tickLblSkip',
	'tickMarkSkip',
	'noMultiLvlLbl',
	'baseTimeUnit',
	'majorTimeUnit',
	'minorTimeUnit',
	'extLst',
];
const SERIES_AFTER_SP_PR = [
	'invertIfNegative',
	'pictureOptions',
	'marker',
	'dPt',
	'dLbls',
	'trendline',
	'errBars',
	'cat',
	'val',
	'xVal',
	'yVal',
	'smooth',
	'shape',
	'explosion',
	'bubbleSize',
	'bubble3D',
	'extLst',
];
const DLBLS_AFTER_SP_PR = [
	'txPr',
	'dLblPos',
	'showLegendKey',
	'showVal',
	'showCatName',
	'showSerName',
	'showPercent',
	'showBubbleSize',
	'separator',
	'showLeaderLines',
	'leaderLines',
	'extLst',
];
const AXIS_TAGS = ['catAx', 'valAx', 'dateAx', 'serAx'];

function keyOf(node: XmlObject, name: string, getLocalName: GetLocalName): string | undefined {
	return Object.keys(node).find((key) => !key.startsWith('@_') && getLocalName(key) === name);
}

function child(node: XmlObject | undefined, name: string, getLocalName: GetLocalName) {
	if (!node) {
		return undefined;
	}
	const key = keyOf(node, name, getLocalName);
	const value = key ? node[key] : undefined;
	return value && typeof value === 'object' && !Array.isArray(value)
		? (value as XmlObject)
		: undefined;
}

function children(node: XmlObject, name: string, getLocalName: GetLocalName): XmlObject[] {
	const key = keyOf(node, name, getLocalName);
	const value = key ? node[key] : undefined;
	if (Array.isArray(value)) {
		return value.filter((v): v is XmlObject => Boolean(v) && typeof v === 'object');
	}
	return value && typeof value === 'object' ? [value as XmlObject] : [];
}

/** Apply `format` to `owner/c:spPr`, inserting a new one before the first `after` child. */
export function applyFormatToOwner(
	owner: XmlObject,
	format: PptxChartElementFormat,
	after: readonly string[],
	getLocalName: GetLocalName,
): void {
	const spPrKey = keyOf(owner, 'spPr', getLocalName);
	const existing = spPrKey ? (owner[spPrKey] as XmlObject | undefined) : undefined;
	const next = applyChartElementFormatToSpPr(
		existing && typeof existing === 'object' ? existing : undefined,
		format,
		getLocalName,
	);
	if (spPrKey) {
		if (next) {
			owner[spPrKey] = next;
		} else {
			delete owner[spPrKey];
		}
		return;
	}
	if (!next) {
		return;
	}
	const entries = Object.entries(owner);
	const at = entries.findIndex(
		([key]) => !key.startsWith('@_') && after.includes(getLocalName(key)),
	);
	entries.splice(at < 0 ? entries.length : at, 0, ['c:spPr', next]);
	for (const key of Object.keys(owner)) {
		delete owner[key];
	}
	for (const [key, value] of entries) {
		owner[key] = value;
	}
}

/** The axis node in `plotArea` the model's `axes[axisIndex]` describes. */
function findAxisNode(
	plotArea: XmlObject,
	chartData: PptxChartData,
	axisIndex: number,
	getLocalName: GetLocalName,
): XmlObject | undefined {
	const axis = chartData.axes?.[axisIndex];
	if (!axis) {
		return undefined;
	}
	const nodes = AXIS_TAGS.flatMap((tag) =>
		children(plotArea, tag, getLocalName).map((node) => ({ tag, node })),
	);
	if (axis.axisId !== undefined) {
		const byId = nodes.find(
			({ node }) => child(node, 'axId', getLocalName)?.['@_val'] === String(axis.axisId),
		);
		if (byId) {
			return byId.node;
		}
	}
	const sameType = (chartData.axes ?? [])
		.slice(0, axisIndex)
		.filter((a) => a.axisType === axis.axisType).length;
	return nodes.filter(({ tag }) => tag === axis.axisType)[sameType]?.node;
}

function formatEntries(chartData: PptxChartData) {
	return Object.entries(chartData.elementFormats ?? {}).flatMap(([key, format]) => {
		const target = parseChartFormatTargetKey(key);
		return target ? [{ target, format }] : [];
	});
}

/**
 * Series-phase writes: series formats and one series' data labels. Call while
 * the model's series still line up index for index with `seriesNodes` (before
 * a combo chart is split back into per-type containers).
 */
export function applySeriesElementFormatsToXml(
	seriesNodes: readonly XmlObject[],
	chartData: PptxChartData,
	getLocalName: GetLocalName,
): void {
	for (const { target, format } of formatEntries(chartData)) {
		if (
			target.kind !== 'series' &&
			!(target.kind === 'dataLabels' && target.seriesIndex !== undefined)
		) {
			continue;
		}
		const node = seriesNodes[target.seriesIndex ?? -1];
		if (!node) {
			continue;
		}
		if (target.kind === 'series') {
			applyFormatToOwner(node, format, SERIES_AFTER_SP_PR, getLocalName);
		} else {
			const dLbls = child(node, 'dLbls', getLocalName);
			if (dLbls) {
				applyFormatToOwner(dLbls, format, DLBLS_AFTER_SP_PR, getLocalName);
			}
		}
	}
}

/** Every `c:dLbls` of the plot area: each chart group's own and each series'. */
function allDataLabelNodes(plotArea: XmlObject, getLocalName: GetLocalName): XmlObject[] {
	const out: XmlObject[] = [];
	for (const key of Object.keys(plotArea)) {
		if (!getLocalName(key).endsWith('Chart')) {
			continue;
		}
		const container = plotArea[key];
		if (!container || typeof container !== 'object' || Array.isArray(container)) {
			continue;
		}
		const owners = [
			container as XmlObject,
			...children(container as XmlObject, 'ser', getLocalName),
		];
		for (const owner of owners) {
			const dLbls = child(owner, 'dLbls', getLocalName);
			if (dLbls) {
				out.push(dLbls);
			}
		}
	}
	return out;
}

/**
 * Chrome-phase writes: chart area, plot area, title, legend, axes and
 * gridlines. Call after the other chart-space serializers so the recorded
 * format is the last word on those `c:spPr` nodes.
 */
export function applyChromeElementFormatsToXml(
	chartSpace: XmlObject,
	chartData: PptxChartData,
	getLocalName: GetLocalName,
): void {
	const chartRoot = child(chartSpace, 'chart', getLocalName);
	const plotArea = child(chartRoot, 'plotArea', getLocalName);
	for (const { target, format } of formatEntries(chartData)) {
		switch (target.kind) {
			case 'chartArea':
				applyFormatToOwner(chartSpace, format, CHART_SPACE_AFTER_SP_PR, getLocalName);
				break;
			case 'plotArea':
				if (plotArea) {
					applyFormatToOwner(plotArea, format, ['extLst'], getLocalName);
				}
				break;
			case 'title':
			case 'legend': {
				const owner = child(chartRoot, target.kind, getLocalName);
				if (owner) {
					applyFormatToOwner(owner, format, TEXT_OWNER_AFTER_SP_PR, getLocalName);
				}
				break;
			}
			case 'axis':
			case 'majorGridlines':
			case 'minorGridlines': {
				const axis = plotArea
					? findAxisNode(plotArea, chartData, target.axisIndex, getLocalName)
					: undefined;
				if (!axis) {
					break;
				}
				if (target.kind === 'axis') {
					applyFormatToOwner(axis, format, AXIS_AFTER_SP_PR, getLocalName);
				} else {
					const lines = child(axis, target.kind, getLocalName);
					if (lines) {
						applyFormatToOwner(lines, format, ['extLst'], getLocalName);
					}
				}
				break;
			}
			case 'dataLabels':
				if (plotArea && target.seriesIndex === undefined) {
					for (const dLbls of allDataLabelNodes(plotArea, getLocalName)) {
						applyFormatToOwner(dLbls, format, DLBLS_AFTER_SP_PR, getLocalName);
					}
				}
				break;
			default:
				break;
		}
	}
}
