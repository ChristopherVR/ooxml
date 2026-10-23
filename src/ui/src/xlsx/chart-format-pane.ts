// Excel's chart Format pane: the shared `office-ui-chart-format-pane` over the chart's `ChartSpace`
// (`xlsxChartFormatSource`, the neutral `chartSpaceFormatSource`). Its typed edit intents go to the
// edit session's `editChart`, which applies them with `applyChartEditIntent` as one undo step; the
// Fill section of the chart area, plot area, title, legend and series is Excel's own fill controls,
// slotted, which edit through `updateChart`.
import { parseChartFormatTargetKey, type ChartFormatTarget } from 'ooxml-core/chart';
import { chartSpaceOf, type ChartFillPart } from 'ooxml-core/xlsx';
import { activeChart, editing, icon, type Command, type EditorContext } from 'ooxml-core/xlsx/ui';
import { chartView, createRefEvaluator } from 'ooxml-core/xlsx';
import {
	defineChartFormatPane,
	type ChartFormatEditEvent,
	type ChartFormatTargetEvent,
	type OfficeUiChartFormatPane,
} from '../chart/format-pane';
import { backgroundFillBinding } from './chart-fill-binding';
import {
	formatsSeries,
	workbookChartTheme,
	xlsxChartFormatSource,
	xlsxFillsTarget,
	XLSX_CHART_FILL_SLOT,
} from './chart-format-source';
import { createSeriesFill } from './chart-series-fill';

export function chartAreaCommand(): Command {
	return editing({
		id: 'chart.format-area',
		label: 'Format Chart Area',
		icon: icon('chartColumn'),
		lock: 'objects',
		enabled: (ctx) => !!activeChart(ctx),
		run: (ctx, arg) => void ctx.dialogs.open('format-chart-area', arg),
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
			return !!chart && formatsSeries(chart);
		},
		run: (ctx, arg) => void ctx.dialogs.open('format-chart-series', arg),
	});
}

const BACKGROUNDS: readonly string[] = ['chartArea', 'plotArea', 'title', 'legend'];

/** Shared captions carry `{{name}}`; Excel's dictionary is keyed by English text with `{name}`. */
const englishKey = (fallback: string) => fallback.replace(/\{\{(\w+)\}\}/gu, '{$1}');

/** The whole-number range of the spacing fields, for the message a refused value gets. */
const SPACING: Record<string, [number, number]> = {
	'series.gapWidth': [0, 500],
	'series.overlap': [-100, 100],
	'series.firstSliceAngle': [0, 360],
	'series.holeSize': [10, 90],
};

export function createChartFormatPane(ctx: EditorContext) {
	const doc = ctx.host.ownerDocument;
	defineChartFormatPane(doc.defaultView?.customElements);
	const element = doc.createElement('office-ui-chart-format-pane') as OfficeUiChartFormatPane;
	element.className = 'xve-chart-series-pane';
	element.hidden = true;
	let target = 'chartArea';
	let seriesIndex = 0;
	let part: ChartFillPart = 'chartArea';
	const seriesFill = createSeriesFill(ctx, () => seriesIndex);
	const backgroundFill = createSeriesFill(
		ctx,
		() => 0,
		backgroundFillBinding(() => part),
	);
	for (const fill of [seriesFill, backgroundFill]) {
		fill.element.slot = XLSX_CHART_FILL_SLOT;
		// The pane's own section toggle names the Fill section.
		fill.element.querySelector('h3')?.setAttribute('hidden', '');
	}
	const parsed = (): ChartFormatTarget | undefined => parseChartFormatTargetKey(target);
	const isSeries = () => parsed()?.kind === 'series';
	const commandId = () => (isSeries() ? 'chart.format-series' : 'chart.format-area');
	const close = () => {
		seriesFill.refresh(undefined, undefined);
		backgroundFill.refresh(undefined, undefined);
		element.hidden = true;
		ctx.grid()?.focus();
	};
	const refresh = () => {
		if (element.hidden) return;
		const found = activeChart(ctx),
			book = ctx.workbook();
		if (!found || !book) {
			close();
			return;
		}
		const view = chartView(
			book,
			ctx.activeSheet(),
			found.chart,
			createRefEvaluator(book, ctx.session()?.calc, { sheet: ctx.activeSheet() }),
		);
		const space = chartSpaceOf(book, found.chart);
		const theme = workbookChartTheme(book);
		const listed = xlsxChartFormatSource(found.chart, space, theme, undefined);
		if (!listed.targets.some((option) => option.key === target)) target = 'chartArea';
		const selected = parsed();
		seriesIndex = selected?.kind === 'series' ? selected.seriesIndex : 0;
		part = BACKGROUNDS.includes(target) ? (target as ChartFillPart) : 'chartArea';
		const slotted = selected !== undefined && xlsxFillsTarget(selected);
		const [active, idle] = isSeries() ? [seriesFill, backgroundFill] : [backgroundFill, seriesFill];
		idle.refresh(undefined, undefined);
		idle.element.remove();
		if (slotted) {
			if (active.element.parentNode !== element) element.append(active.element);
			active.refresh(found.chart, view);
		} else {
			active.refresh(undefined, undefined);
			active.element.remove();
		}
		element.source = xlsxChartFormatSource(found.chart, space, theme, selected);
		element.target = target;
		element.disabled = !ctx.commands.isEnabled(commandId());
	};
	element.addEventListener('office-chart-format-target', (event) => {
		target = (event as ChartFormatTargetEvent).detail.target;
		refresh();
	});
	element.addEventListener('office-chart-format-close', close);
	element.addEventListener('office-chart-format-edit', (event) => {
		const { field, intent } = (event as ChartFormatEditEvent).detail;
		if (!ctx.commands.isEnabled(commandId())) return refresh();
		const range = SPACING[field];
		const found = activeChart(ctx);
		if (!intent || !found) {
			if (range)
				ctx.toast(
					ctx.t('Enter a whole number from {min} to {max}.', { min: range[0], max: range[1] }),
					'warning',
				);
			return refresh();
		}
		try {
			ctx.session()?.editChart(ctx.activeSheet(), found.index, intent);
		} catch (error) {
			if (!(error instanceof RangeError)) throw error;
			ctx.toast(ctx.t('This chart element cannot take that setting.'), 'warning');
		}
		refresh();
	});
	const relocalize = () => {
		// The pane's own dictionary translates its captions; Excel's wording wins where it has one.
		element.locale = ctx.locale();
		element.translateLabel = (key, params, fallback) => {
			const english = englishKey(fallback ?? key);
			const own = ctx.t(english, params as Record<string, string | number> | undefined);
			return own === english ? key : own;
		};
		refresh();
	};
	const open = (focus: string) => {
		if (!ctx.commands.isEnabled(commandId())) return;
		element.hidden = false;
		refresh();
		element.shadowRoot?.querySelector<HTMLElement>(focus)?.focus();
	};
	ctx.dialogs.register('format-chart-area', async (_ctx, arg) => {
		target = typeof arg === 'string' && BACKGROUNDS.includes(arg) ? arg : 'chartArea';
		open('[data-chart-format-target]');
	});
	ctx.dialogs.register('format-chart-series', async (_ctx, arg) => {
		const found = activeChart(ctx);
		const index =
			typeof arg === 'number' && Number.isInteger(arg) && found?.chart.series[arg] ? arg : 0;
		target = `series:${index}`;
		open('[data-chart-format-field="series.overlap"] input');
	});
	relocalize();
	return { element, refresh, relocalize, close };
}
