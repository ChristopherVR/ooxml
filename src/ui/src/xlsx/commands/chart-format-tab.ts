// Excel's chart Format contextual tab: Current Selection (Chart Elements, Format Selection, Reset to
// Match Style), Shape Styles, WordArt Styles and Size act on the chart element picked in Chart
// Elements, through the shared chart edit intents (`ooxml-core/chart`) the Format pane also uses.
// Insert Shapes and Arrange are the Shape Format tab's own commands. The pane stays for detailed
// options: Format Selection opens it on the picked element.
import { chartFormatTargetKey, parseChartFormatTargetKey } from 'ooxml-core/chart';
import { chartSpaceOf } from 'ooxml-core/xlsx';
import { activeChart, editing, type Command, type EditorContext } from 'ooxml-core/xlsx/ui';
import { translateChartFormatText } from '../../chart/format-pane-common';
import { chartSpaceTargetOptions } from '../../chart/format-pane-space-targets';

/** The picked element per editor, remembered for the chart (sheet and drawing) it was picked on. */
const picked = new WeakMap<EditorContext, { sheet: number; index: number; key: string }>();

/** Shared captions carry `{{name}}`; Excel's dictionary is keyed by English text with `{name}`. */
const englishKey = (fallback: string) => fallback.replace(/\{\{(\w+)\}\}/gu, '{$1}');

function selected(ctx: EditorContext) {
	const found = activeChart(ctx);
	const book = ctx.workbook();
	if (!found || !book || found.chart.editBlocked) return undefined;
	return { ...found, space: chartSpaceOf(book, found.chart), sheet: ctx.activeSheet() };
}

/** The Chart Elements entries of the selected chart, captioned in the display language. */
export function chartElementOptions(ctx: EditorContext) {
	const chart = selected(ctx);
	if (!chart) return [];
	const translate = (
		key: string,
		params?: Readonly<Record<string, string | number>>,
		fallback?: string,
	) => ctx.t(englishKey(fallback ?? key), params as Record<string, string | number> | undefined);
	return chartSpaceTargetOptions(chart.space)
		.filter((option) => option.target.kind !== 'title' || chart.chart.title)
		.map((option) => ({
			value: option.key,
			label: translateChartFormatText(option.label, translate),
		}));
}

/** The element the tab formats: the one picked in Chart Elements, else the chart area. */
export function chartFormatTarget(ctx: EditorContext) {
	const chart = selected(ctx);
	if (!chart) return undefined;
	const memo = picked.get(ctx);
	const keys = chartElementOptions(ctx).map((option) => option.value);
	const key =
		memo && memo.sheet === chart.sheet && memo.index === chart.index && keys.includes(memo.key)
			? memo.key
			: 'chartArea';
	const target = parseChartFormatTargetKey(key);
	return target && { ...chart, key, target };
}

export type ChartFormatTabTarget = NonNullable<ReturnType<typeof chartFormatTarget>>;

const reason = 'Select an editable chart.';

export const chartFormatActions = {
	'chart.format-selection': 'Format Selection',
	'chart.reset-style': 'Reset to Match Style',
	'chart.shape-style': 'Shape Styles',
	'chart.shape-fill': 'Shape Fill',
	'chart.shape-outline': 'Shape Outline',
	'chart.shape-effects': 'Shape Effects',
	'chart.word-art-style': 'WordArt Styles',
	'chart.text-fill': 'Text Fill',
	'chart.text-outline': 'Text Outline',
	'chart.size': 'Size',
} as const;

export type ChartFormatAction = keyof typeof chartFormatActions;

export function chartFormatTabCommands(): Command[] {
	return [
		editing({
			id: 'chart.element',
			label: 'Chart Elements',
			unavailable: reason,
			lock: 'objects',
			enabled: (ctx) => !!chartFormatTarget(ctx),
			value: (ctx) => chartFormatTarget(ctx)?.key ?? 'chartArea',
			run: (ctx, arg) => {
				const chart = chartFormatTarget(ctx);
				if (!chart || typeof arg !== 'string' || !parseChartFormatTargetKey(arg)) return;
				const key = chartFormatTargetKey(parseChartFormatTargetKey(arg)!);
				picked.set(ctx, { sheet: chart.sheet, index: chart.index, key });
				ctx.requestRender();
			},
		}),
		...Object.entries(chartFormatActions).map(([id, label]) =>
			editing({
				id,
				label,
				unavailable: reason,
				lock: 'objects',
				enabled: (ctx) => !!chartFormatTarget(ctx),
				run: async (ctx, arg) => {
					const chart = chartFormatTarget(ctx);
					if (!chart) return;
					const { runChartFormat } = await import('./chart-format-tab-actions');
					await runChartFormat(ctx, chart, id as ChartFormatAction, arg);
				},
			}),
		),
	];
}
