import { chartPaletteSeriesColorChoice, findChartColorPalette } from '../../chart/color-palettes';
import {
	chartColorStyleId,
	chartColorStyleXml,
	CHART_COLOR_STYLE_CONTENT_TYPE,
	CHART_COLOR_STYLE_REL,
} from '../../chart/color-style';
import { parseDrawingColorIn } from '../../diagram/drawing-color';
import { drawingColorXml } from '../../diagram/write-color';
import { drawingFillXml } from '../../diagram/write-fill';
import type { DiagramColor } from '../../diagram/types';
import { NS, children, elements, first, parseXml, type XmlElement } from '../../xml/index';
import { sameChartColor } from '../edit/chart-colors';
import type { ChartObject, ChartSeries } from '../model';
import { THEME_SLOTS } from '../layout/colors';
import { RelationshipSet, type PackageWriter } from './package-writer';

type Doc = ReturnType<typeof parseXml>;
const FILLS = new Set(['solidFill', 'gradFill', 'pattFill', 'blipFill', 'grpFill', 'noFill']);

/** Primary color of a series, preserving a manual fill before palette defaults. */
export function chartSeriesFill(chart: ChartObject, series: ChartSeries, index: number): string {
	const imported = series.fill && drawingFillXml(series.fill);
	if (imported) return imported;
	const palette = findChartColorPalette(chart.colorPalette ?? 10)!;
	const choice =
		series.drawingColor ??
		(series.color?.rgb
			? { kind: 'srgb' as const, value: series.color.rgb.slice(-6), transforms: [] }
			: undefined) ??
		(series.color?.theme !== undefined
			? {
					kind: 'scheme' as const,
					value: THEME_SLOTS[series.color.theme] ?? 'accent1',
					transforms: [],
				}
			: undefined) ??
		chartPaletteSeriesColorChoice(
			palette ?? findChartColorPalette(10)!,
			index,
			chart.series.length,
		);
	return `<a:solidFill>${drawingColorXml(choice)}</a:solidFill>`;
}

function setFill(doc: Doc, parent: XmlElement, xml: string): void {
	const wrapper = parseXml(`<w xmlns:a="${NS.a}">${xml}</w>`);
	const fill = elements(wrapper.documentElement)[0]!;
	const existing = elements(parent).filter(
		(node) => node.namespaceURI === NS.a && FILLS.has(node.localName),
	);
	const next = doc.importNode(fill, true);
	if (existing[0]) parent.replaceChild(next, existing[0]);
	else
		parent.insertBefore(
			next,
			first(parent, 'ln', NS.a) ?? first(parent, 'effectLst', NS.a) ?? null,
		);
	for (const extra of existing.slice(1)) parent.removeChild(extra);
}

function recolorMarker(
	doc: Doc,
	marker: XmlElement | undefined,
	previous: DiagramColor,
	next: DiagramColor,
): void {
	for (const paint of [marker, first(marker, 'ln', NS.a)]) {
		if (paint && sameChartColor(parseDrawingColorIn(first(paint, 'solidFill', NS.a)), previous))
			setFill(doc, paint, `<a:solidFill>${drawingColorXml(next)}</a:solidFill>`);
	}
}

/** Recolors primary automatic series/marker fills without replacing axes, labels or effects. */
export function patchChartColors(
	doc: Doc,
	chartNode: XmlElement,
	model: ChartObject,
	before: ChartObject,
): void {
	const plotArea = first(chartNode, 'plotArea', NS.c);
	if (!plotArea) return;
	const plots = elements(plotArea);
	const plot = plots.find((node) => children(node, 'ser', NS.c).length);
	if (!plot) return;
	const previousPalette = findChartColorPalette(before.colorPalette ?? 10);
	const palette = findChartColorPalette(model.colorPalette ?? 10);
	const lineLike = ['line', 'scatter', 'radar'].includes(model.chartType);
	const markerOnly =
		model.chartType === 'scatter' &&
		first(plot, 'scatterStyle', NS.c)?.getAttribute('val') === 'marker';
	children(plot, 'ser', NS.c).forEach((ser, index) => {
		const series = model.series[index];
		if (!series) return;
		const old = before.series[index];
		if (
			!sameChartColor(old?.drawingColor, series.drawingColor) ||
			JSON.stringify(old?.fill) !== JSON.stringify(series.fill) ||
			JSON.stringify(old?.color) !== JSON.stringify(series.color)
		) {
			let spPr = markerOnly
				? first(first(ser, 'marker', NS.c), 'spPr', NS.c)
				: first(ser, 'spPr', NS.c);
			if (!spPr) {
				spPr = doc.createElementNS(NS.c, 'c:spPr');
				if (markerOnly) {
					let marker = first(ser, 'marker', NS.c);
					if (!marker) {
						marker = doc.createElementNS(NS.c, 'c:marker');
						ser.insertBefore(marker, first(ser, 'dPt', NS.c) ?? first(ser, 'xVal', NS.c) ?? null);
					}
					marker.insertBefore(spPr, first(marker, 'extLst', NS.c) ?? null);
				} else
					ser.insertBefore(
						spPr,
						first(ser, 'marker', NS.c) ??
							first(ser, 'invertIfNegative', NS.c) ??
							first(ser, 'dPt', NS.c) ??
							first(ser, 'cat', NS.c) ??
							first(ser, 'xVal', NS.c) ??
							first(ser, 'val', NS.c) ??
							null,
					);
			}
			let paint = lineLike && !markerOnly ? first(spPr, 'ln', NS.a) : spPr;
			if (!paint) {
				paint = doc.createElementNS(NS.a, 'a:ln');
				spPr.appendChild(paint);
			}
			setFill(doc, paint, chartSeriesFill(model, series, index));
		}
		if (previousPalette && palette && before.colorPalette !== model.colorPalette) {
			const marker = first(first(ser, 'marker', NS.c), 'spPr', NS.c);
			const previous = chartPaletteSeriesColorChoice(previousPalette, index, model.series.length);
			const next = chartPaletteSeriesColorChoice(palette, index, model.series.length);
			recolorMarker(doc, marker, previous, next);
			for (const point of children(ser, 'dPt', NS.c))
				recolorMarker(doc, first(first(point, 'marker', NS.c), 'spPr', NS.c), previous, next);
		}
		if (
			JSON.stringify([old?.pointColors, old?.pointFills]) !==
			JSON.stringify([series.pointColors, series.pointFills])
		) {
			const points = children(ser, 'dPt', NS.c);
			const indices = new Set([
				...Object.keys(old?.pointColors ?? {}),
				...Object.keys(old?.pointFills ?? {}),
				...Object.keys(series.pointColors ?? {}),
				...Object.keys(series.pointFills ?? {}),
			]);
			for (const key of indices) {
				if (!/^\d+$/.test(key)) continue;
				const pointIndex = Number(key);
				let point = points.find(
					(node) => Number(first(node, 'idx', NS.c)?.getAttribute('val')) === pointIndex,
				);
				const color = series.pointColors?.[pointIndex];
				const fill = series.pointFills?.[pointIndex];
				const xml = fill
					? drawingFillXml(fill)
					: color
						? `<a:solidFill>${drawingColorXml(color)}</a:solidFill>`
						: undefined;
				let spPr = first(point, 'spPr', NS.c);
				if (!xml) {
					for (const paint of spPr ? elements(spPr) : [])
						if (paint.namespaceURI === NS.a && FILLS.has(paint.localName)) spPr!.removeChild(paint);
					continue;
				}
				if (!point) {
					point = doc.createElementNS(NS.c, 'c:dPt');
					const idx = doc.createElementNS(NS.c, 'c:idx');
					idx.setAttribute('val', key);
					point.appendChild(idx);
					ser.insertBefore(
						point,
						first(ser, 'dLbls', NS.c) ??
							first(ser, 'trendline', NS.c) ??
							first(ser, 'errBars', NS.c) ??
							first(ser, 'cat', NS.c) ??
							first(ser, 'xVal', NS.c) ??
							first(ser, 'val', NS.c) ??
							null,
					);
				}
				if (!spPr) {
					spPr = doc.createElementNS(NS.c, 'c:spPr');
					point.insertBefore(
						spPr,
						first(point, 'pictureOptions', NS.c) ?? first(point, 'extLst', NS.c) ?? null,
					);
				}
				setFill(doc, spPr, xml);
			}
		}
	});
}

/** Writes a private color-style copy, so editing one chart cannot recolor another chart. */
export function writeChartColorStyle(
	writer: PackageWriter,
	chartPart: string,
	chart: ChartObject,
): void {
	if (chart.colorPalette === undefined) return;
	const palette = findChartColorPalette(chart.colorPalette);
	if (!palette) return;
	const source = writer.source;
	const existing = [...(source?.rels(chartPart).entries() ?? [])];
	const entry = existing.find(([, rel]) => rel.type.endsWith('/chartColorStyle'));
	const sourcePart = entry && source?.target(chartPart, entry[1]);
	const original = sourcePart && source?.text(sourcePart);
	if (original && chartColorStyleId(original) === palette.id) return;
	const part = writer.uniqueName(
		(n) => `xl/charts/colors${n}.xml`,
		source ? new Set(source.parts.keys()) : undefined,
	);
	writer.add(
		part,
		chartColorStyleXml(palette, original || undefined),
		CHART_COLOR_STYLE_CONTENT_TYPE,
	);
	const rels = new RelationshipSet();
	for (const [id, rel] of existing) rels.keep(id, rel);
	const target =
		chartPart.split('/').slice(0, -1).join('/') === 'xl/charts'
			? part.slice('xl/charts/'.length)
			: '/' + part;
	if (entry) rels.keep(entry[0], { ...entry[1], target });
	else rels.add(CHART_COLOR_STYLE_REL, target);
	writer.rels(chartPart, rels);
}
