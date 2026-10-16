/**
 * Pure serialization helper for writing a chart's chart-level data labels
 * (`c:dLbls` directly under each chart-type container, applying to every
 * series) back into the parsed chart XML tree on save.
 *
 * Dependency-light (only a `getLocalName` resolver) so it works for both
 * prefixed (`c:dLbls`) and namespace-stripped (`dLbls`) keys and can be
 * unit-tested without a full save round-trip.
 *
 * @module utils/chart-data-labels-serializer
 */

import type { PptxChartDataLabelsRange, XmlObject } from '../types';
import { applyDataLabelsRangeToDLbls } from './chart-data-labels-range';

/** Resolve a possibly-prefixed XML key to its local name. */
type GetLocalName = (key: string) => string;

/** The data-label-relevant subset of `PptxChartStyle`. */
export interface ChartDataLabelStyle {
	hasDataLabels?: boolean;
	dataLabels?: {
		showValue?: boolean;
		showCategory?: boolean;
		showSeriesName?: boolean;
		showPercent?: boolean;
		showLegendKey?: boolean;
		showBubbleSize?: boolean;
		separator?: string;
		showLeaderLines?: boolean;
		position?: string;
		/** Label-specific number format (`c:dLbls/c:numFmt/@formatCode`), distinct from the series' own cell format (C2-G16). */
		numberFormat?: string;
		/** PowerPoint 2013+ "Value From Cells" (`c15:datalabelsRange`); see `chart-data-labels-range.ts`. */
		dataLabelsRange?: PptxChartDataLabelsRange;
		/** `c15:showDataLabelsRange` at the group level, paired with {@link dataLabelsRange}. */
		showDataLabelsRange?: boolean;
	};
}

const POSITION_VALUES = new Set([
	'bestFit',
	'b',
	'ctr',
	'inBase',
	'inEnd',
	'l',
	'outEnd',
	'r',
	't',
]);
const CHILD_ORDER = [
	'dLbl',
	'delete',
	'numFmt',
	'spPr',
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
] as const;
const MODELED = new Set([
	'delete',
	'dLblPos',
	'showLegendKey',
	'showVal',
	'showCatName',
	'showSerName',
	'showPercent',
	'showBubbleSize',
	'separator',
	'showLeaderLines',
]);

const SHOW_FLAGS = [
	'showValue',
	'showCategory',
	'showSeriesName',
	'showPercent',
	'showLegendKey',
	'showBubbleSize',
] as const;
const SHOW_TAGS: Record<(typeof SHOW_FLAGS)[number], string> = {
	showValue: 'showVal',
	showCategory: 'showCatName',
	showSeriesName: 'showSerName',
	showPercent: 'showPercent',
	showLegendKey: 'showLegendKey',
	showBubbleSize: 'showBubbleSize',
};
/** `c:ser` children that follow `c:dLbls` (CT_BarSer / CT_LineSer / CT_PieSer / ...). */
const SERIES_AFTER_DLBLS = [
	'trendline',
	'errBars',
	'cat',
	'val',
	'xVal',
	'yVal',
	'smooth',
	'shape',
	'bubbleSize',
	'bubble3D',
	'extLst',
];

function findKey(obj: XmlObject, local: string, getLocalName: GetLocalName): string | undefined {
	return Object.keys(obj).find((k) => getLocalName(k) === local);
}

function boolVal(on: boolean | undefined): XmlObject {
	return { '@_val': on ? '1' : '0' };
}

function ordered(
	existing: XmlObject | undefined,
	modeled: XmlObject,
	getLocalName: GetLocalName,
): XmlObject {
	const entries = Object.entries(existing ?? {}).filter(([key]) => !MODELED.has(getLocalName(key)));
	entries.push(...Object.entries(modeled));
	entries.sort(([a], [b]) => {
		const rank = (key: string) => {
			const index = CHILD_ORDER.indexOf(getLocalName(key) as (typeof CHILD_ORDER)[number]);
			return index < 0 ? CHILD_ORDER.length - 1 : index;
		};
		return rank(a) - rank(b);
	});
	return Object.fromEntries(entries) as XmlObject;
}

/** Insert `c:dLbls` after the last `c:ser` child (schema order), preserving key order. */
function insertAfterLastSeries(
	container: XmlObject,
	dLbls: XmlObject,
	getLocalName: GetLocalName,
): void {
	const keys = Object.keys(container);
	let lastSer = -1;
	keys.forEach((k, i) => {
		if (getLocalName(k) === 'ser') {
			lastSer = i;
		}
	});
	const entries = keys.map((k) => [k, container[k]] as const);
	const at = lastSer === -1 ? entries.length : lastSer + 1;
	entries.splice(at, 0, ['c:dLbls', dLbls] as const);
	for (const k of keys) {
		delete container[k];
	}
	for (const [k, v] of entries) {
		container[k] = v;
	}
}

/**
 * Build a `c:dLbls` element from the requested options. Preserves an existing
 * node's `numFmt`/`spPr`/`txPr` styling (in schema order) and then writes the
 * `dLblPos` and `show*` flags in schema order.
 */
function buildDLbls(
	existing: XmlObject | undefined,
	opts: NonNullable<ChartDataLabelStyle['dataLabels']>,
	getLocalName: GetLocalName,
): XmlObject {
	const built: XmlObject = {};
	if (opts.numberFormat !== undefined) {
		built['c:numFmt'] = { '@_formatCode': opts.numberFormat, '@_sourceLinked': '0' };
	}
	if (opts.position) {
		if (!POSITION_VALUES.has(opts.position)) {
			throw new RangeError(`Invalid data label position: ${opts.position}`);
		}
		built['c:dLblPos'] = { '@_val': opts.position };
	}
	built['c:showLegendKey'] = boolVal(opts.showLegendKey);
	built['c:showVal'] = boolVal(opts.showValue);
	built['c:showCatName'] = boolVal(opts.showCategory);
	built['c:showSerName'] = boolVal(opts.showSeriesName);
	built['c:showPercent'] = boolVal(opts.showPercent);
	built['c:showBubbleSize'] = boolVal(opts.showBubbleSize);
	if (opts.separator !== undefined) {
		built['c:separator'] = opts.separator;
	}
	if (opts.showLeaderLines !== undefined) {
		built['c:showLeaderLines'] = boolVal(opts.showLeaderLines);
	}
	// numFmt joins spPr/txPr in the "preserve unless the model overrides it"
	// group (unlike dLblPos/separator/etc., which MODELED always rewrites):
	// an untouched chart's label number format must survive editing its show
	// flags. Only strip the OLD node from the passthrough when a fresh one is
	// about to replace it, so the two don't both end up in the merged result.
	const baseExisting =
		opts.numberFormat === undefined || !existing
			? existing
			: (Object.fromEntries(
					Object.entries(existing).filter(([key]) => getLocalName(key) !== 'numFmt'),
				) as XmlObject);
	return ordered(baseExisting, built, getLocalName);
}

/**
 * Apply chart-level data-label visibility/content onto the plot area.
 *
 * - `style.hasDataLabels === true` writes a `c:dLbls` under every chart-type
 *   container with the requested `show*` flags and optional `dLblPos`,
 *   defaulting to showing the value when no content flag is set.
 * - `style.hasDataLabels === false` disables labels via `<c:dLbls><c:delete
 *   val="1"/></c:dLbls>`.
 * - `undefined` leaves the chart untouched so unedited charts round-trip via
 *   the original XML.
 *
 * Mutates `plotArea` in place.
 */
export function applyChartDataLabelsToXml(
	plotArea: XmlObject,
	style: ChartDataLabelStyle,
	getLocalName: GetLocalName,
): void {
	if (style.hasDataLabels === undefined) {
		return;
	}

	const chartTypeKeys = Object.keys(plotArea).filter((k) => getLocalName(k).endsWith('Chart'));
	// Two groups of one type (primary and secondary axis) are one key holding an array.
	const containers = chartTypeKeys.flatMap((key) =>
		([] as Array<XmlObject | undefined>).concat(
			plotArea[key] as XmlObject | XmlObject[] | undefined,
		),
	);
	for (const container of containers) {
		if (!container || typeof container !== 'object') {
			continue;
		}
		const existingKey = findKey(container, 'dLbls', getLocalName);

		if (style.hasDataLabels === false) {
			const existing = existingKey ? (container[existingKey] as XmlObject) : undefined;
			const off = ordered(existing, { 'c:delete': { '@_val': '1' } }, getLocalName);
			if (existingKey) {
				container[existingKey] = off;
			} else {
				insertAfterLastSeries(container, off, getLocalName);
			}
			continue;
		}

		const opts = style.dataLabels ?? {};
		// Fall back to the value only when no content flag is declared at all: an
		// explicit all-off group (PowerPoint writes one on every chart, and the
		// pane writes one when only some series show labels) stays off.
		const anyDeclared = SHOW_FLAGS.some((flag) => opts[flag] !== undefined);
		const effective = anyDeclared ? opts : { ...opts, showValue: true };

		const existing = existingKey ? (container[existingKey] as XmlObject) : undefined;
		const built = buildDLbls(existing, effective, getLocalName);
		applyDataLabelsRangeToDLbls(
			built,
			opts.dataLabelsRange,
			opts.showDataLabelsRange,
			getLocalName,
		);

		if (existingKey) {
			container[existingKey] = built;
		} else {
			insertAfterLastSeries(container, built, getLocalName);
		}
	}
}

function flagOf(node: XmlObject | undefined, tag: string, getLocalName: GetLocalName) {
	const key = node ? findKey(node, tag, getLocalName) : undefined;
	const value = key ? (node![key] as XmlObject | undefined)?.['@_val'] : undefined;
	return value === undefined ? undefined : value === '1' || value === 'true';
}

/**
 * Write one series' own data labels (`c:ser/c:dLbls`) from its
 * `dataLabelOptions`, which override the chart-level group for that series
 * (Format Data Labels on one series). Only written when a declared show flag
 * or the position differs from what the series node authors, so an unedited
 * series keeps its labels exactly as written. Mutates `seriesNode`.
 */
export function applySeriesDataLabelOptionsToXml(
	seriesNode: XmlObject,
	opts: NonNullable<ChartDataLabelStyle['dataLabels']> | undefined,
	getLocalName: GetLocalName,
): void {
	if (!opts) {
		return;
	}
	const existingKey = findKey(seriesNode, 'dLbls', getLocalName);
	const existing = existingKey ? (seriesNode[existingKey] as XmlObject) : undefined;
	const deleted = flagOf(existing, 'delete', getLocalName) === true;
	const changed =
		SHOW_FLAGS.some(
			(flag) =>
				opts[flag] !== undefined &&
				(deleted
					? opts[flag]
					: opts[flag] !== Boolean(flagOf(existing, SHOW_TAGS[flag], getLocalName))),
		) ||
		(opts.position !== undefined &&
			opts.position !==
				(existing
					? (existing[findKey(existing, 'dLblPos', getLocalName) ?? ''] as XmlObject | undefined)?.[
							'@_val'
						]
					: undefined));
	if (!changed) {
		return;
	}
	const built = buildDLbls(existing, opts, getLocalName);
	if (existingKey) {
		seriesNode[existingKey] = built;
		return;
	}
	const entries = Object.entries(seriesNode);
	const at = entries.findIndex(
		([key]) => !key.startsWith('@_') && SERIES_AFTER_DLBLS.includes(getLocalName(key)),
	);
	entries.splice(at < 0 ? entries.length : at, 0, ['c:dLbls', built]);
	for (const key of Object.keys(seriesNode)) {
		delete seriesNode[key];
	}
	for (const [key, value] of entries) {
		seriesNode[key] = value;
	}
}
