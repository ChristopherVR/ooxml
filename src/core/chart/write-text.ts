// `c:chartSpace` writer: data sources with their caches, chart text (`c:tx`), titles and data
// labels, each in the element order of ECMA-376 Part 1, 21.2.
import type {
	ChartDataCache,
	ChartDataLabel,
	ChartDataLabelOptions,
	ChartDataLabels,
	ChartDataSource,
	ChartText,
	ChartTitle,
} from './model-series';
import { layoutXml, numberFormatXml, shapePropertiesXml, textBodyXml } from './write-shape';
import { writeTextBody } from './write-text-body';
import {
	elementXml,
	escapeAttribute,
	escapeText,
	raw,
	valXml,
	type ChartWriteContext,
} from './write-util';

/** Text content, written as an open and close pair even when empty (as the xlsx writer did). */
const textElement = (local: string, value: string) =>
	`<c:${local}>${escapeText(value)}</c:${local}>`;

/** The points of a cache; a point without an index takes its position (the index is required). */
function cachePoints(cache: ChartDataCache): string {
	return cache.points
		.map((point, position) => {
			const format =
				point.formatCode === undefined ? '' : ` formatCode="${escapeAttribute(point.formatCode)}"`;
			return `<c:pt idx="${point.index ?? position}"${format}>${textElement('v', point.value)}</c:pt>`;
		})
		.join('');
}

function cacheXml(local: string, cache: ChartDataCache | undefined, withFormat: boolean): string {
	if (!cache) return '';
	const format =
		withFormat && cache.formatCode !== undefined ? textElement('formatCode', cache.formatCode) : '';
	return elementXml(local, format + valXml('ptCount', cache.pointCount) + cachePoints(cache));
}

/** The reference or literal inside a data source element. */
function dataSourceInner(source: ChartDataSource): string {
	const formula = textElement('f', source.formula ?? '');
	switch (source.kind) {
		case 'numRef':
			return `<c:numRef>${formula}${cacheXml('numCache', source.cache, true)}</c:numRef>`;
		case 'strRef':
			return `<c:strRef>${formula}${cacheXml('strCache', source.cache, false)}</c:strRef>`;
		case 'multiLvlStrRef': {
			const levels = source.levels
				? elementXml(
						'multiLvlStrCache',
						valXml('ptCount', source.levelPointCount) +
							source.levels.map((level) => elementXml('lvl', cachePoints(level))).join(''),
					)
				: '';
			return `<c:multiLvlStrRef>${formula}${levels}</c:multiLvlStrRef>`;
		}
		case 'numLit':
			return cacheXml('numLit', source.cache ?? { type: 'number', points: [] }, true);
		case 'strLit':
			return cacheXml('strLit', source.cache ?? { type: 'string', points: [] }, false);
		default:
			return '';
	}
}

/** A data source element (`c:cat`, `c:val`, `c:xVal`...). */
export const dataSourceXml = (local: string, source: ChartDataSource | undefined): string =>
	source ? elementXml(local, dataSourceInner(source)) : '';

/**
 * Chart text (`c:tx`). Series names (`seriesName`) take a reference or a `c:v` value; titles and
 * labels take a reference or rich text. Text given only as `text` is written as a value or as one
 * rich paragraph.
 */
export function chartTextXml(
	context: ChartWriteContext,
	text: ChartText | undefined,
	seriesName: boolean,
): string {
	if (!text) return '';
	if (text.reference) return elementXml('tx', dataSourceInner(text.reference));
	if (seriesName) {
		const value = text.value ?? text.text;
		return elementXml('tx', value === undefined ? '' : textElement('v', value));
	}
	if (text.rich) return elementXml('tx', textBodyXml(context, text.rich, 'rich'));
	if (text.text === undefined) return '<c:tx/>';
	const paragraphs = text.text.split('\n').map((line) => ({ runs: [{ text: line }] }));
	return `<c:tx><c:rich>${writeTextBody({ paragraphs })}</c:rich></c:tx>`;
}

/** A chart or axis title (`c:title`). */
export function titleXml(context: ChartWriteContext, title: ChartTitle | undefined): string {
	if (!title) return '';
	const tx = title.tx ?? (title.text === undefined ? undefined : { text: title.text });
	return elementXml(
		'title',
		chartTextXml(context, tx, false) +
			layoutXml(context, title.layout) +
			valXml('overlay', title.overlay) +
			shapePropertiesXml(context, title.spPr) +
			textBodyXml(context, title.txPr) +
			raw(context, title.extLst),
	);
}

function labelOptionsXml(context: ChartWriteContext, options: ChartDataLabelOptions): string {
	return (
		numberFormatXml(options.numberFormat) +
		shapePropertiesXml(context, options.spPr) +
		textBodyXml(context, options.txPr) +
		valXml('dLblPos', options.position) +
		valXml('showLegendKey', options.showLegendKey) +
		valXml('showVal', options.showValue) +
		valXml('showCatName', options.showCategoryName) +
		valXml('showSerName', options.showSeriesName) +
		valXml('showPercent', options.showPercent) +
		valXml('showBubbleSize', options.showBubbleSize) +
		(options.separator === undefined ? '' : textElement('separator', options.separator))
	);
}

function labelXml(context: ChartWriteContext, label: ChartDataLabel): string {
	return elementXml(
		'dLbl',
		valXml('idx', label.index) +
			valXml('delete', label.deleted) +
			layoutXml(context, label.layout) +
			chartTextXml(context, label.tx, false) +
			labelOptionsXml(context, label) +
			raw(context, label.extLst),
	);
}

/** Data labels of a group or series (`c:dLbls`). */
export function dataLabelsXml(
	context: ChartWriteContext,
	labels: ChartDataLabels | undefined,
): string {
	if (!labels) return '';
	const leaderLines = labels.leaderLines
		? elementXml('leaderLines', shapePropertiesXml(context, labels.leaderLines.spPr))
		: '';
	return elementXml(
		'dLbls',
		labels.labels.map((label) => labelXml(context, label)).join('') +
			valXml('delete', labels.deleted) +
			labelOptionsXml(context, labels) +
			valXml('showLeaderLines', labels.showLeaderLines) +
			leaderLines +
			raw(context, labels.extLst),
	);
}
