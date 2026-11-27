// What the chart Format tab's commands do, loaded on first use. Every edit is one undo step of
// shared chart edit intents on the element picked in Chart Elements.
import {
	chartFillIntent,
	chartOutlineIntent,
	chartResetFormatIntents,
	chartShadowIntent,
	chartShapeFormatIntents,
	chartTargetHasText,
	chartTextFormatIntents,
	type ChartEditIntent,
} from 'ooxml-core/chart';
import { parseShapeStylePick, WORD_ART_STYLES, type DrawingColor } from 'ooxml-core/drawingml';
import {
	anchorToPixelBox,
	chartDrawingColor,
	chartShapeStyleFormat,
	createGridMetrics,
	pixelBoxToAnchor,
} from 'ooxml-core/xlsx';
import { pickColor, type EditorContext } from 'ooxml-core/xlsx/ui';
import { field, numberInput } from '../dialogs/fields';
import { showDialog } from '../dialogs/frame';
import { bounded } from './drawing-format-actions';
import type { ChartFormatAction, ChartFormatTabTarget } from './chart-format-tab';
import { chartFormatTarget } from './chart-format-tab';

const rgb = (value: string): DrawingColor => ({ kind: 'srgb', value, transforms: [] });

/** A colour control's pick: a colour, or null for No Fill / No Outline. */
function picked(arg: unknown): DrawingColor | null | undefined {
	const color = pickColor(arg);
	if (!color) return null;
	return chartDrawingColor(color);
}

function apply(
	ctx: EditorContext,
	chart: ChartFormatTabTarget,
	label: string,
	intents: ChartEditIntent[],
): void {
	const session = ctx.session();
	const live = chartFormatTarget(ctx);
	if (!session || !intents.length || live?.index !== chart.index || live.key !== chart.key) return;
	try {
		session.batch(label, () => {
			for (const intent of intents) session.editChart(chart.sheet, chart.index, intent);
		});
	} catch (error) {
		if (!(error instanceof RangeError)) throw error;
		ctx.toast(ctx.t('This chart element cannot take that setting.'), 'warning');
	}
}

async function resize(ctx: EditorContext, chart: ChartFormatTabTarget, label: string) {
	const sheet = ctx.workbook()?.sheets[chart.sheet];
	if (!sheet) return;
	const metrics = createGridMetrics(sheet, { zoom: 100 });
	const box = anchorToPixelBox(sheet, chart.chart.anchor, metrics);
	const width = numberInput(ctx, Math.round(box.w), 4, 10000);
	const height = numberInput(ctx, Math.round(box.h), 4, 10000);
	await showDialog(ctx, {
		name: 'chart-size',
		heading: label,
		body: [field(ctx, 'Width (px)', width), field(ctx, 'Height (px)', height)],
		submit: () => {
			const w = bounded(width);
			const h = bounded(height);
			const live = chartFormatTarget(ctx);
			if (w === undefined || h === undefined || live?.index !== chart.index) return undefined;
			ctx
				.session()
				?.setDrawingAnchor(
					chart.sheet,
					chart.index,
					pixelBoxToAnchor(sheet, { ...box, w, h }, metrics, chart.chart.anchor),
				);
			return true;
		},
	});
}

export async function runChartFormat(
	ctx: EditorContext,
	chart: ChartFormatTabTarget,
	id: ChartFormatAction,
	arg: unknown,
): Promise<void> {
	const label = ctx.t(ctx.commands.get(id)?.label ?? '');
	const { target } = chart;
	switch (id) {
		case 'chart.format-selection':
			await ctx.dialogs.open('format-chart-selection', chart.key);
			return;
		case 'chart.reset-style':
			return apply(ctx, chart, label, chartResetFormatIntents(chart.space, target));
		case 'chart.shape-style': {
			const pick = typeof arg === 'string' ? parseShapeStylePick(arg) : undefined;
			const book = ctx.workbook();
			const format = pick && book && chartShapeStyleFormat(book, pick);
			if (format) apply(ctx, chart, label, chartShapeFormatIntents(target, format));
			return;
		}
		case 'chart.shape-fill':
		case 'chart.shape-outline': {
			const color = picked(arg);
			if (color === undefined) return;
			const intent =
				id === 'chart.shape-fill'
					? chartFillIntent(target, color)
					: chartOutlineIntent(target, color);
			return apply(ctx, chart, label, [intent]);
		}
		case 'chart.shape-effects':
			if (typeof arg === 'boolean') apply(ctx, chart, label, [chartShadowIntent(target, arg)]);
			return;
		case 'chart.word-art-style':
		case 'chart.text-fill':
		case 'chart.text-outline': {
			// The chart area's text is all the chart's text, as Excel applies it.
			if (target.kind !== 'chartArea' && !chartTargetHasText(target)) {
				ctx.toast(ctx.t('This chart element cannot take that setting.'), 'warning');
				return;
			}
			if (id === 'chart.word-art-style') {
				const style = typeof arg === 'string' ? WORD_ART_STYLES[Number(arg)] : undefined;
				if (!style || !/^\d{1,2}$/.test(String(arg))) return;
				return apply(
					ctx,
					chart,
					label,
					chartTextFormatIntents(target, {
						fill: rgb(style.fill),
						outline: style.outline ? rgb(style.outline) : null,
						shadow: !!style.shadow,
					}),
				);
			}
			const color = picked(arg);
			if (color === undefined) return;
			if (id === 'chart.text-outline')
				return apply(ctx, chart, label, chartTextFormatIntents(target, { outline: color }));
			// Text keeps a fill: No Fill is not offered for chart text.
			if (color) apply(ctx, chart, label, chartTextFormatIntents(target, { fill: color }));
			return;
		}
		case 'chart.size':
			await resize(ctx, chart, label);
	}
}
