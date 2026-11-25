import { activeChart, editing, icon, type Command, type EditorContext } from 'ooxml-core/xlsx/ui';

export function activeChartEx(ctx: EditorContext) {
	const index = ctx.selection.get().drawing;
	const book = ctx.workbook();
	if (index === undefined || !book) return undefined;
	const drawing = book.sheets[ctx.activeSheet()]?.drawings[index];
	return drawing?.kind === 'unsupported' &&
		drawing.chartExSpace &&
		drawing.chartExPartName &&
		!drawing.chartExIssues?.length
		? { drawing, index }
		: undefined;
}

export function chartAreaCommand(): Command {
	return editing({
		id: 'chart.format-area',
		label: 'Format Chart Area',
		icon: icon('chartColumn'),
		lock: 'objects',
		enabled: (ctx) => {
			const chart = activeChart(ctx)?.chart;
			return (!!chart && !chart.editBlocked) || !!activeChartEx(ctx);
		},
		run: async (ctx, arg) => {
			await ctx.dialogs.open(activeChartEx(ctx) ? 'chart-ex' : 'format-chart-area', arg);
		},
	});
}

export function chartSeriesCommand(): Command {
	return editing({
		id: 'chart.format-series',
		label: 'Format Data Series',
		icon: icon('chartColumn'),
		lock: 'objects',
		enabled: (ctx) => {
			const chart = activeChart(ctx)?.chart;
			return !!chart && !chart.editBlocked && chart.series.length > 0;
		},
		run: async (ctx, arg) => {
			await ctx.dialogs.open('format-chart-series', arg);
		},
	});
}
