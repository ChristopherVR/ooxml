// Insert Chart: a chart type gallery, grouping, title and a live preview drawn by the core's
// chartView + renderChartSvg. OK adds a ChartObject over the selection (or, from Chart Design,
// changes the active chart's type). The extended families (ids `cx:<family>`) are offered by Insert
// Chart only; OK adds an extended chart over the selection through the session's `addChartEx`.
import { type ChartObject, type ChartType, chartView, renderChartSvg } from 'ooxml-core/xlsx';
import { activeChart, editChart } from 'ooxml-core/xlsx/ui';
import { CHART_TYPES } from '../commands/insert';
import { regionOf, target } from 'ooxml-core/xlsx/ui';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { showDialog } from './frame';
import { text } from './fields';
import { buildChart, buildChartEx, evaluateRef, INSERT_CHART_EX_TYPES } from 'ooxml-core/xlsx/ui';
import {
	chartExSummaryFromModel,
	renderChartSummarySvg,
	type ChartExFamily,
} from 'ooxml-core/chart';
import { defineChartPicker } from '../../dialog/chart-picker';
import type { OfficeChartValue } from '../../dialog/chart-picker';

type Picker = HTMLElement & {
	types: { id: string; label: string }[];
	groupings: { id: string; label: string }[];
	groupedTypes: string[];
	labels: { chartType: string; grouping: string; title: string; preview: string };
	drawPreview: (
		value: OfficeChartValue,
		size: { width: number; height: number },
		tile: boolean,
	) => string;
	value: OfficeChartValue;
	changeOnly: boolean;
};

export interface InsertChartProps {
	type?: ChartType;
	change?: boolean;
}

type Grouping = NonNullable<ChartObject['grouping']>;
const MAP_NOTICE =
	'Map outlines here are approximate. Office draws a new map only after it adds map data through its online service.';
const GROUPINGS: ReadonlyArray<readonly [Grouping, string]> = [
	['clustered', 'Clustered'],
	['stacked', 'Stacked'],
	['percentStacked', '100% Stacked'],
];

export function openInsertChart(
	ctx: EditorContext,
	props: InsertChartProps = {},
): Promise<ChartObject | undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const change = props.change === true ? activeChart(ctx) : undefined;
	const range = regionOf(t);
	const doc = ctx.host.ownerDocument;
	defineChartPicker(doc.defaultView?.customElements);
	const picker = doc.createElement('office-ui-chart-picker') as Picker;
	picker.types = [
		...CHART_TYPES.map(([id, label]) => ({ id: id as string, label: ctx.t(label) })),
		...(change
			? []
			: INSERT_CHART_EX_TYPES.map(([family, label]) => ({
					id: `cx:${family}`,
					label: ctx.t(label),
				}))),
	];
	const familyOf = (type: string): ChartExFamily | undefined =>
		INSERT_CHART_EX_TYPES.find(([family]) => `cx:${family}` === type)?.[0];
	const mapNotice = ctx.t(MAP_NOTICE);
	const extended = (value: OfficeChartValue, family: ChartExFamily) =>
		buildChartEx(t.workbook, t.sheet, range, family, {
			...(value.title.trim() ? { title: value.title.trim() } : {}),
		});
	picker.groupings = GROUPINGS.map(([id, label]) => ({ id, label: ctx.t(label) }));
	picker.groupedTypes = ['column', 'bar'];
	picker.labels = {
		chartType: ctx.t('Chart type'),
		grouping: ctx.t('Grouping'),
		title: ctx.t('Chart title'),
		preview: ctx.t('Preview'),
	};
	picker.changeOnly = change !== undefined;
	picker.value = {
		type: props.type ?? change?.chart.chartType ?? 'column',
		grouping: change?.chart.grouping ?? 'clustered',
		title: change?.chart.title ?? '',
	};
	const current = (value: OfficeChartValue): ChartObject => {
		const grouping = value.grouping as Grouping;
		if (change)
			return { ...structuredClone(change.chart), chartType: value.type as ChartType, grouping };
		return buildChart(t.workbook, t.sheet, range, value.type as ChartType, {
			grouping,
			...(value.title.trim() ? { title: value.title.trim() } : {}),
		});
	};
	picker.drawPreview = (value, size, tile) => {
		const family = familyOf(value.type);
		if (family) {
			try {
				const summary = chartExSummaryFromModel(extended(value, family).chartExSpace).summary;
				if (!summary) return '';
				const { title: _title, ...bare } = summary;
				return renderChartSummarySvg(tile ? { ...bare, showLegend: false } : summary, {
					width: size.width,
					height: size.height,
					mapApproximationLabel: mapNotice,
				});
			} catch {
				return '';
			}
		}
		// A type tile is drawn bare (no title or legend); the large preview shows the chart as it will be.
		const chart = { ...current(value), ...(tile ? { title: '', showLegend: false } : {}) };
		return renderChartSvg(
			chartView(t.workbook, t.sheet, chart, (ref) => evaluateRef(t.workbook, t.sheet, ref)),
			size.width,
			size.height,
		);
	};
	// A new region map has no geography cache: say what the preview's outlines are.
	const note = doc.createElement('p');
	note.setAttribute('role', 'note');
	const updateNote = () => {
		note.textContent = familyOf(picker.value.type) === 'regionMap' ? mapNotice : '';
	};
	picker.addEventListener('office-chart-change', updateNote);
	updateNote();
	const body: HTMLElement[] = [picker, note];
	if (!change && t.ws.rows.size === 0)
		body.push(text(ctx, 'Select the data for the chart first, then insert it.'));
	return showDialog<ChartObject>(ctx, {
		name: 'insert-chart',
		heading: change ? 'Change Chart Type' : 'Insert Chart',
		wide: true,
		body,
		opened: () =>
			picker.shadowRoot
				?.querySelector<HTMLButtonElement>(`[data-gallery-item="${picker.value.type}"]`)
				?.focus(),
		submit: () => {
			const value = picker.value;
			const family = familyOf(value.type);
			if (family) {
				const built = extended(value, family);
				ctx.selection.set({ drawing: t.session.addChartEx(t.sheet, built) });
				return built.chartEx!;
			}
			const chart = current(value);
			if (change)
				editChart(ctx, () => ({
					chartType: value.type as ChartType,
					...(value.type === 'column' || value.type === 'bar'
						? { grouping: value.grouping as Grouping }
						: {}),
				}));
			else {
				const { kind: _kind, ...model } = chart;
				ctx.selection.set({ drawing: t.session.addChart(t.sheet, model) });
			}
			return chart;
		},
	});
}
