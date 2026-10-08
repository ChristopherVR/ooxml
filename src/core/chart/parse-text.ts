// `c:chartSpace` parser: data sources and their caches, chart text, titles and data labels.
import { NS, children, type XmlElement } from '../xml/index';
import { chartCacheValues } from './data-cache';
import type {
	ChartCachePoint,
	ChartDataCache,
	ChartDataLabel,
	ChartDataLabelOptions,
	ChartDataLabels,
	ChartDataSource,
	ChartDataSourceKind,
	ChartText,
	ChartTitle,
} from './model-series';
import {
	assignDefined,
	attribute,
	cBool,
	chartLines,
	chartTextBody,
	cChild,
	cNumber,
	cVal,
	extensionList,
	manualLayout,
	numberFormat,
	reportUnmodelled,
	shapeProperties,
	textProperties,
	type ChartParseContext,
} from './parse-util';

function parseCache(node: XmlElement): ChartDataCache {
	const cache: ChartDataCache = {
		type: node.localName === 'numCache' || node.localName === 'numLit' ? 'number' : 'string',
		points: children(node, 'pt', NS.c).map((pt) => {
			const point: ChartCachePoint = { value: cChild(pt, 'v')?.textContent ?? '' };
			const index = Number(attribute(pt, 'idx') ?? Number.NaN);
			if (Number.isInteger(index) && index >= 0) point.index = index;
			const formatCode = attribute(pt, 'formatCode');
			if (formatCode !== undefined) point.formatCode = formatCode;
			return point;
		}),
	};
	const count = Number(cVal(node, 'ptCount') ?? Number.NaN);
	if (Number.isInteger(count) && count >= 0) cache.pointCount = count;
	const formatCode = cChild(node, 'formatCode')?.textContent;
	if (formatCode) cache.formatCode = formatCode;
	return cache;
}

/** A data source element (`c:cat`, `c:val`, `c:tx`...): reference and cache, or literal values. */
export function parseDataSource(element: XmlElement | undefined): ChartDataSource | undefined {
	if (!element) return undefined;
	const source: ChartDataSource = {};
	const ref =
		cChild(element, 'numRef') ?? cChild(element, 'strRef') ?? cChild(element, 'multiLvlStrRef');
	if (ref) {
		source.kind = ref.localName as ChartDataSourceKind;
		const formula = cChild(ref, 'f')?.textContent;
		if (formula) source.formula = formula;
		const cache = cChild(ref, 'numCache') ?? cChild(ref, 'strCache');
		if (cache) source.cache = parseCache(cache);
		const levels = cChild(ref, 'multiLvlStrCache');
		if (levels) {
			source.levels = children(levels, 'lvl', NS.c).map(parseCache);
			const count = Number(cVal(levels, 'ptCount') ?? Number.NaN);
			if (Number.isInteger(count) && count >= 0) source.levelPointCount = count;
		}
		return source;
	}
	const literal = cChild(element, 'numLit') ?? cChild(element, 'strLit');
	if (literal) {
		source.kind = literal.localName as ChartDataSourceKind;
		source.cache = parseCache(literal);
	}
	return source;
}

/** Chart text (`c:tx`); `c:v` is read only for series names (`allowValue`). */
export function parseChartText(
	context: ChartParseContext,
	tx: XmlElement | undefined,
	allowValue: boolean,
): ChartText | undefined {
	if (!tx) return undefined;
	const text: ChartText = {};
	const rich = chartTextBody(cChild(tx, 'rich'));
	if (rich) text.rich = rich;
	if (cChild(tx, 'strRef')) {
		const reference = parseDataSource(tx);
		if (reference) text.reference = reference;
	}
	const value = allowValue ? cChild(tx, 'v')?.textContent : undefined;
	if (value !== undefined && value !== null) text.value = value;
	const cached = chartCacheValues(text.reference?.cache)[0];
	const flat = text.value ?? (rich ? rich.text : cached === undefined ? undefined : String(cached));
	if (flat !== undefined) text.text = flat;
	reportUnmodelled(context, tx, new Set(allowValue ? ['strRef', 'v'] : ['rich', 'strRef']));
	return text;
}

const TITLE = new Set(['tx', 'layout', 'overlay', 'spPr', 'txPr', 'extLst']);

/** A chart or axis title (`c:title`). */
export function parseTitle(context: ChartParseContext, element: XmlElement): ChartTitle {
	const tx = parseChartText(context, cChild(element, 'tx'), false);
	reportUnmodelled(context, element, TITLE);
	return assignDefined<ChartTitle>(
		{},
		{
			tx,
			text: tx?.text,
			overlay: cBool(context, element, 'overlay'),
			layout: manualLayout(element),
			spPr: shapeProperties(element),
			txPr: textProperties(element),
			extLst: extensionList(element),
		},
	);
}

const LABEL_OPTIONS = [
	'delete',
	'showLegendKey',
	'showVal',
	'showCatName',
	'showSerName',
	'showPercent',
	'showBubbleSize',
	'separator',
	'dLblPos',
	'numFmt',
	'spPr',
	'txPr',
	'extLst',
];

function labelOptions(context: ChartParseContext, element: XmlElement): ChartDataLabelOptions {
	const flag = (local: string) => cBool(context, element, local);
	return assignDefined<ChartDataLabelOptions>(
		{},
		{
			deleted: flag('delete'),
			showLegendKey: flag('showLegendKey'),
			showValue: flag('showVal'),
			showCategoryName: flag('showCatName'),
			showSeriesName: flag('showSerName'),
			showPercent: flag('showPercent'),
			showBubbleSize: flag('showBubbleSize'),
			separator: cChild(element, 'separator')?.textContent ?? undefined,
			position: cVal(element, 'dLblPos'),
			numberFormat: numberFormat(element),
			spPr: shapeProperties(element),
			txPr: textProperties(element),
			extLst: extensionList(element),
		},
	);
}

const LABEL = new Set([...LABEL_OPTIONS, 'idx', 'layout', 'tx']);
const LABELS = new Set([...LABEL_OPTIONS, 'dLbl', 'showLeaderLines', 'leaderLines']);

/** Data labels (`c:dLbls`) of a group or series. */
export function parseDataLabels(
	context: ChartParseContext,
	element: XmlElement | undefined,
): ChartDataLabels | undefined {
	if (!element) return undefined;
	const labels = children(element, 'dLbl', NS.c).map((node) => {
		reportUnmodelled(context, node, LABEL);
		return assignDefined<ChartDataLabel>(labelOptions(context, node), {
			index: cNumber(context, node, 'idx'),
			tx: parseChartText(context, cChild(node, 'tx'), false),
			layout: manualLayout(node),
		});
	});
	reportUnmodelled(context, element, LABELS);
	return assignDefined<ChartDataLabels>(
		{ ...labelOptions(context, element), labels },
		{
			showLeaderLines: cBool(context, element, 'showLeaderLines'),
			leaderLines: chartLines(element, 'leaderLines'),
		},
	);
}
