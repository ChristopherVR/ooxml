import {
	chartSeriesGradientPatch,
	chartView,
	createRefEvaluator,
	type ChartObject,
	type ChartViewModel,
} from 'ooxml-core/xlsx';
import { drawingColorCss } from 'ooxml-core/diagram';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import type { createGradientStopTrack } from '../form/gradient-stop-track';
import { createChartGradientPreview } from './chart-gradient-preview';

export type GradientParameter = 'position' | 'brightness' | 'transparency';

/** Core computes the temporary gradient; only UI paint changes until the range commits. */
export function seriesRangePreview(
	ctx: EditorContext,
	chart: ChartObject,
	view: ChartViewModel,
	series: number,
	stop: number,
	drawing: number,
	track: ReturnType<typeof createGradientStopTrack>,
	color: HTMLButtonElement,
	valid: () => boolean,
) {
	const gradient = view.series[series]?.gradient;
	const book = ctx.workbook();
	const sheet = ctx.activeSheet();
	let preview: ReturnType<typeof createChartGradientPreview> | undefined;
	const strip = (fill: NonNullable<typeof gradient>) => {
		track.preview(
			fill.stops.map((item) => ({
				position: item.position,
				color: drawingColorCss({ hex: item.color, alpha: item.opacity ?? 1, unapplied: [] })!,
			})),
		);
		color.style.setProperty('--series-fill', fill.stops[stop]?.color ?? 'transparent');
	};
	return (property: GradientParameter, value: number | undefined) => {
		if (value === undefined) {
			preview?.restore();
			preview = undefined;
			if (valid() && gradient) strip(gradient);
			return;
		}
		if (
			!valid() ||
			!book ||
			!gradient ||
			ctx.workbook() !== book ||
			ctx.activeSheet() !== sheet ||
			!ctx.commands.isEnabled('chart.format-series')
		)
			return;
		const result = chartSeriesGradientPatch(chart, series, {
			kind: 'stop',
			index: stop,
			[property]: value,
		});
		const next = result
			? chartView(
					book,
					sheet,
					{ ...chart, ...result.patch },
					createRefEvaluator(book, ctx.session()?.calc, { sheet }),
				).series[series]?.gradient
			: gradient;
		if (!next) return;
		preview ??= createChartGradientPreview(ctx.root, drawing, series, gradient);
		preview.paint(next);
		strip(next);
	};
}
