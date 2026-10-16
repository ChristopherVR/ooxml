/**
 * Draws Text Options (`c:txPr` fill, outline, shadow and glow, custom angle
 * and direction) on the chart texts: the title, legend entries, axis labels
 * and data labels, each falling back to the chart area's text options.
 *
 * A solid text colour already reaches the view model through the flat text
 * fields (the Format operation mirrors it); this pass adds what those cannot
 * carry: transparency, no fill, gradient fill, outline, shadow and glow, as
 * text drawn into an effect image under the real text nodes (which are kept,
 * transparent where the image draws their body, for selection and
 * hit-testing), and a custom angle or direction on axis labels.
 *
 * Text in the effect image uses the browser's installed fonts: a web font
 * the page loaded is not visible inside an SVG image, so a chart whose text
 * uses one draws its gradient, outline, shadow or glow in a fallback face.
 *
 * @module render/chart-format/chart-format-render-text
 */
import type { PptxChartData, PptxChartElementText } from 'ooxml-core/pptx';

import { computeChartLegendLayout } from '../chart-legend-layout';
import type { ChartViewModel, LegendEntry, SvgText } from '../chart-view-model-types';
import { opacityOf, ptToPx, safeColor } from './chart-format-effects-svg';
import { chartAxisIndex, chartTextBox } from './chart-format-geometry';
import { gradientPaint } from './chart-format-paint-svg';

const r2 = (v: number) => Math.round(v * 100) / 100;
const esc = (value: string) =>
	value
		.replace(/&/gu, '&amp;')
		.replace(/</gu, '&lt;')
		.replace(/>/gu, '&gt;')
		.replace(/"/gu, '&quot;');

let serial = 0;
const uid = (prefix: string) => `${prefix}t${++serial}`;

/** The text options of `key` over the chart area's (chart-wide) ones. */
export function effectiveTextFormat(
	data: PptxChartData,
	key: string,
): PptxChartElementText | undefined {
	const formats = data.elementFormats;
	const own = formats?.[key];
	const base = formats?.chartArea;
	const pick = (record: typeof own) => (record && !record.reset ? record.text : undefined);
	const ownText = pick(own);
	const baseText = key === 'chartArea' ? undefined : pick(base);
	if (!ownText && !baseText) return undefined;
	return {
		...baseText,
		...ownText,
		...(baseText?.effects || ownText?.effects
			? { effects: { ...baseText?.effects, ...ownText?.effects } }
			: {}),
	};
}

/** Whether the text body (fill or outline) must be drawn in the effect image. */
function drawsBody(text: PptxChartElementText): boolean {
	return text.fill?.kind === 'gradient' || text.outline?.kind === 'solid';
}

function hasEffects(text: PptxChartElementText): boolean {
	return Boolean(text.effects?.shadow || (text.effects?.glow && text.effects.glow.radius > 0));
}

/** A drawable text: the SVG text node and its paint. */
interface TextRun {
	node: SvgText;
	fill: string;
}

function textAttrs(node: SvgText): string {
	return [
		`x="${r2(node.x)}"`,
		`y="${r2(node.y)}"`,
		`font-size="${r2(node.fontSize)}"`,
		node.fontFamily
			? `font-family="${esc(node.fontFamily)}"`
			: 'font-family="Calibri, Arial, sans-serif"',
		node.fontWeight ? `font-weight="${node.fontWeight}"` : '',
		node.fontStyle ? `font-style="${node.fontStyle}"` : '',
		`text-anchor="${node.textAnchor}"`,
		node.dominantBaseline ? `dominant-baseline="${node.dominantBaseline}"` : '',
		node.transform ? `transform="${esc(node.transform)}"` : '',
	]
		.filter(Boolean)
		.join(' ');
}

/** Shadow, glow and (when it needs the image) the body of `runs`, as effect-image markup. */
function textMarkup(runs: readonly TextRun[], text: PptxChartElementText): string {
	let out = '';
	const shadow = text.effects?.shadow ?? undefined;
	const glow = text.effects?.glow && text.effects.glow.radius > 0 ? text.effects.glow : undefined;
	for (const { node, fill } of runs) {
		const label = esc(node.text);
		const attrs = textAttrs(node);
		if (shadow) {
			const id = uid('s');
			const dist = ptToPx(shadow.distance);
			const rad = (shadow.angle * Math.PI) / 180;
			out += `<filter id="${id}" x="-50%" y="-50%" width="200%" height="200%"><feGaussianBlur stdDeviation="${r2(ptToPx(shadow.blur) / 2)}"/></filter><g filter="url(#${id})" opacity="${r2(opacityOf(shadow.transparency))}" transform="translate(${r2(dist * Math.cos(rad))} ${r2(dist * Math.sin(rad))})"><text ${attrs} fill="${safeColor(shadow.color)}">${label}</text></g>`;
		}
		if (glow) {
			const id = uid('g');
			const radius = ptToPx(glow.radius);
			const color = safeColor(glow.color, '#FFFFFF');
			out += `<filter id="${id}" x="-50%" y="-50%" width="200%" height="200%"><feGaussianBlur stdDeviation="${r2(radius / 3)}"/></filter><g filter="url(#${id})" opacity="${r2(opacityOf(glow.transparency))}"><text ${attrs} fill="${color}" stroke="${color}" stroke-width="${r2(radius)}" stroke-linejoin="round">${label}</text></g>`;
		}
		if (drawsBody(text)) {
			const box = chartTextBox(node, 0);
			const paint =
				text.fill?.kind === 'gradient' && text.fill.gradient
					? gradientPaint(text.fill.gradient, {
							x: box.left,
							y: box.top,
							w: box.right - box.left,
							h: box.bottom - box.top,
						})
					: undefined;
			const bodyFill =
				text.fill?.kind === 'none'
					? 'none'
					: paint
						? paint.paint
						: text.fill?.kind === 'solid'
							? safeColor(text.fill.color, fill)
							: fill;
			const fillOpacity =
				text.fill?.kind === 'solid'
					? ` fill-opacity="${r2(opacityOf(text.fill.transparency))}"`
					: '';
			const outline =
				text.outline?.kind === 'solid'
					? ` stroke="${safeColor(text.outline.color)}" stroke-width="${r2(ptToPx(text.outline.width ?? 0.75))}" stroke-opacity="${r2(opacityOf(text.outline.transparency))}" stroke-linejoin="round"`
					: '';
			out += `${paint?.defs ?? ''}<text ${attrs} fill="${bodyFill}"${fillOpacity}${outline}>${label}</text>`;
		}
	}
	return out;
}

/** How a text node looks once `text` is applied (the body moves to the image, or fades). */
function restyle(node: SvgText, text: PptxChartElementText): SvgText {
	let next = node;
	if (drawsBody(text)) {
		next = { ...next, opacity: 0 };
	} else if (text.fill?.kind === 'none') {
		next = { ...next, opacity: 0 };
	} else if (text.fill?.kind === 'solid' && (text.fill.transparency ?? 0) > 0) {
		next = { ...next, opacity: (node.opacity ?? 1) * opacityOf(text.fill.transparency) };
	}
	return next;
}

/** The rotation a custom angle or text direction gives axis label `node`, or the node unchanged. */
function rotated(node: SvgText, text: PptxChartElementText, vertical: boolean): SvgText {
	const angle =
		text.direction === 'vert'
			? 90
			: text.direction === 'vert270'
				? -90
				: text.rotation !== undefined && text.rotation !== null && text.rotation !== 0
					? text.rotation
					: undefined;
	if (angle === undefined) return node;
	// A horizontal axis hangs rotated labels from their end (negative angles) or start.
	const anchor = vertical ? node.textAnchor : angle < 0 ? 'end' : 'start';
	return {
		...node,
		textAnchor: anchor,
		transform: `rotate(${r2(angle)}, ${r2(node.x)}, ${r2(node.y)})`,
	};
}

/** Apply the Text Options of `data` to `vm`; returns the effect-image markup for the texts. */
export function applyTextFormats(
	vm: ChartViewModel,
	data: PptxChartData,
): { vm: ChartViewModel; markup: string } {
	if (!data.elementFormats) return { vm, markup: '' };
	serial = 0;
	let next = vm;
	let markup = '';
	const axisLists: Array<
		[number, 'axisLabels' | 'categoryLabels' | 'secondaryAxisLabels', boolean]
	> = [
		[chartAxisIndex(data, ['catAx', 'dateAx']), 'categoryLabels', data.barDirection === 'bar'],
		[chartAxisIndex(data, ['valAx']), 'axisLabels', data.barDirection !== 'bar'],
		[chartAxisIndex(data, ['valAx'], 1), 'secondaryAxisLabels', true],
	];
	for (const [index, field, vertical] of axisLists) {
		const texts = next[field];
		const text = index >= 0 ? effectiveTextFormat(data, `axis:${index}`) : undefined;
		if (!text || !texts?.length) continue;
		const placed = texts.map((node) => rotated(node, text, vertical));
		if (drawsBody(text) || hasEffects(text)) {
			markup += textMarkup(
				placed.map((node) => ({ node, fill: node.fill })),
				text,
			);
		}
		next = { ...next, [field]: placed.map((node) => restyle(node, text)) };
	}
	const labels = effectiveTextFormat(data, 'dataLabels');
	if (labels && next.dataLabels.length > 0) {
		if (drawsBody(labels) || hasEffects(labels)) {
			markup += textMarkup(
				next.dataLabels.map((node) => ({ node, fill: node.fill })),
				labels,
			);
		}
		next = { ...next, dataLabels: next.dataLabels.map((node) => restyle(node, labels)) };
	}
	const legend = effectiveTextFormat(data, 'legend');
	if (legend && next.legend.length > 0 && data.style?.hasLegend !== false) {
		const items = computeChartLegendLayout(next);
		if (drawsBody(legend) || hasEffects(legend)) {
			markup += textMarkup(
				items.map((item) => ({
					node: {
						kind: 'text',
						x: item.x + 13,
						y: item.y + 3,
						text: item.label,
						fontSize: item.fontSize,
						fill: item.fill,
						textAnchor: 'start',
						...(item.fontFamily ? { fontFamily: item.fontFamily } : {}),
						fontWeight: item.fontWeight === 'bold' ? 'bold' : 'normal',
					},
					fill: item.fill,
				})),
				legend,
			);
		}
		if (drawsBody(legend) || legend.fill?.kind === 'none') {
			next = {
				...next,
				legend: next.legend.map((entry): LegendEntry => ({
					...entry,
					textStyle: { ...entry.textStyle, color: 'transparent' },
				})),
			};
		}
	}
	const title = effectiveTextFormat(data, 'title');
	if (title && next.title) {
		const style = next.titleStyle;
		// A title with typed runs is drawn run by run; its effect text uses the first run's font.
		const run = next.titleRunSpans?.[0];
		const fontSize = run?.fontSize ?? style?.fontSize ?? 12;
		const fontWeight = run?.fontWeight ?? style?.fontWeight ?? 400;
		const fontFamily = run?.fontFamily ?? style?.fontFamily;
		const node: SvgText = {
			kind: 'text',
			x: next.titleX,
			y: next.titleY,
			text: next.titleRunSpans?.length
				? next.titleRunSpans.map((span) => span.text).join('')
				: next.title,
			fontSize,
			fill: run?.fill ?? style?.fill ?? '#000000',
			textAnchor: 'middle',
			fontWeight: fontWeight >= 600 ? 'bold' : 'normal',
			...(fontFamily ? { fontFamily } : {}),
		};
		if (drawsBody(title) || hasEffects(title)) {
			markup += textMarkup([{ node, fill: node.fill }], title);
		}
		if (drawsBody(title) || title.fill?.kind === 'none') {
			next = {
				...next,
				...(style ? { titleStyle: { ...style, fill: 'transparent' } } : {}),
				...(next.titleRunSpans
					? {
							titleRunSpans: next.titleRunSpans.map((span) => ({
								...span,
								fill: 'transparent',
							})),
						}
					: {}),
			};
		}
	}
	return { vm: next, markup };
}
