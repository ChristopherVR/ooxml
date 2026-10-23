// Excel's chart Format pane: the shared `office-ui-chart-format-pane` (Format Chart Area, Plot
// Area, Chart Title, Legend and Data Series in one pane), read through `xlsxChartFormatSource`.
// Its typed edit intents go through the xlsx `ChartPatch` path (`updateChart`, whose paint edits
// call the shared `chart/edit` operations); the Fill section is Excel's own fill controls, slotted.
import type { ChartEditIntent } from 'ooxml-core/chart';
import {
	chartView,
	createRefEvaluator,
	type ChartFillPart,
	type ChartPatch,
} from 'ooxml-core/xlsx';
import { activeChart, editing, icon, type Command, type EditorContext } from 'ooxml-core/xlsx/ui';
import {
	defineChartFormatPane,
	type ChartFormatEditEvent,
	type ChartFormatTargetEvent,
	type OfficeUiChartFormatPane,
} from '../chart/format-pane';
import { backgroundFillBinding } from './chart-fill-binding';
import { formatsSeries, xlsxChartFormatSource, XLSX_CHART_FILL_SLOT } from './chart-format-source';
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

/** The `ChartPatch` of an intent the pane offers in Excel, or undefined for one it does not. */
function patchOf(intent: ChartEditIntent): ChartPatch | undefined {
	if (intent.op === 'groupOptions') {
		const { gapWidth, overlap } = intent.edit;
		if (gapWidth !== undefined) return { barGapWidth: gapWidth };
		if (overlap !== undefined) return { barOverlap: overlap };
	}
	if (intent.op === 'legend' && intent.edit.position)
		return { showLegend: true, legendPosition: intent.edit.position };
	return undefined;
}

const SPACING: Record<string, [number, number]> = {
	'series.gapWidth': [0, 500],
	'series.overlap': [-100, 100],
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
	const isSeries = () => target.startsWith('series:');
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
		const source = xlsxChartFormatSource(found.chart, view);
		if (!source.targets.some((option) => option.key === target)) target = 'chartArea';
		seriesIndex = isSeries() ? Number(target.slice(7)) : 0;
		part = BACKGROUNDS.includes(target) ? (target as ChartFillPart) : 'chartArea';
		const [active, idle] = isSeries() ? [seriesFill, backgroundFill] : [backgroundFill, seriesFill];
		idle.refresh(undefined, undefined);
		idle.element.remove();
		if (active.element.parentNode !== element) element.append(active.element);
		active.refresh(found.chart, view);
		element.source = source;
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
		const patch = intent && patchOf(intent);
		if (!patch) {
			if (range)
				ctx.toast(
					ctx.t('Enter a whole number from {min} to {max}.', { min: range[0], max: range[1] }),
					'warning',
				);
			return refresh();
		}
		const found = activeChart(ctx);
		const changed = Object.entries(patch).some(
			([key, next]) => found?.chart[key as keyof typeof found.chart] !== next,
		);
		if (found && changed) ctx.session()?.updateChart(ctx.activeSheet(), found.index, patch);
		refresh();
	});
	const relocalize = () => {
		element.translateLabel = (_key, params, fallback) =>
			ctx.t(englishKey(fallback ?? _key), params as Record<string, string | number> | undefined);
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
