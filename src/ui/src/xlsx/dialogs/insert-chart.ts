// Insert Chart: a chart type gallery, grouping, title and a live preview drawn by the core's
// chartView + renderChartSvg. OK adds a ChartObject over the selection (or, from Chart Design,
// changes the active chart's type).
import { type ChartObject, type ChartType, chartView, renderChartSvg } from 'ooxml-core/xlsx';
import { activeChart, editChart } from 'ooxml-core/xlsx/ui';
import { CHART_TYPES } from '../commands/insert';
import { regionOf, target } from 'ooxml-core/xlsx/ui';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { el, field, select, text, textInput } from './fields';
import { showDialog } from './frame';
import { buildChart, evaluateRef } from 'ooxml-core/xlsx/ui';
import {
	defineGallery,
	type OfficeUiGallery,
	type OfficeGalleryPickEvent,
} from '../../ribbon/gallery';
import { parseSvgPreview } from '../../ribbon/safe-svg';

export interface InsertChartProps {
	type?: ChartType;
	change?: boolean;
}

type Grouping = NonNullable<ChartObject['grouping']>;
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
	let type: ChartType = props.type ?? change?.chart.chartType ?? 'column';
	const range = regionOf(t);
	defineGallery(ctx.host.ownerDocument.defaultView?.customElements);
	const tiles = ctx.host.ownerDocument.createElement('office-ui-gallery') as OfficeUiGallery;
	tiles.setAttribute('mode', 'panel');
	const grouping = select(ctx, GROUPINGS, change?.chart.grouping ?? 'clustered');
	const groupingField = field(ctx, 'Grouping', grouping);
	const title = textInput(ctx, change?.chart.title ?? '');
	const preview = el(ctx, 'div', 'xve-preview');
	preview.setAttribute('aria-label', ctx.t('Preview'));
	preview.setAttribute('role', 'img');
	const current = (): ChartObject => {
		if (change)
			return {
				...structuredClone(change.chart),
				chartType: type,
				grouping: grouping.value as Grouping,
			};
		return buildChart(t.workbook, t.sheet, range, type, {
			grouping: grouping.value as Grouping,
			...(title.value.trim() ? { title: title.value.trim() } : {}),
		});
	};
	const refresh = (): void => {
		tiles.state = {
			id: 'chart-type',
			label: ctx.t('Chart type'),
			sections: [
				{
					columns: 4,
					tileWidth: 96,
					tileHeight: 60,
					items: CHART_TYPES.map(([key, label]) => ({
						id: key,
						label: ctx.t(label),
						applied: key === type,
						preview: renderChartSvg(
							chartView(
								t.workbook,
								t.sheet,
								{ ...current(), chartType: key, title: '', showLegend: false },
								(ref) => evaluateRef(t.workbook, t.sheet, ref),
							),
							360,
							220,
						),
					})),
				},
			],
		};
		groupingField.hidden = !(type === 'column' || type === 'bar');
		const model = chartView(t.workbook, t.sheet, current(), (ref) =>
			evaluateRef(t.workbook, t.sheet, ref),
		);
		const node = parseSvgPreview(ctx.host.ownerDocument, renderChartSvg(model, 360, 220));
		preview.replaceChildren(...(node ? [node] : []));
	};
	tiles.addEventListener('office-gallery-pick', (event) => {
		const key = (event as OfficeGalleryPickEvent).detail.itemId;
		if (!CHART_TYPES.some(([id]) => id === key)) return;
		type = key as ChartType;
		refresh();
	});
	grouping.addEventListener('change', refresh);
	title.addEventListener('input', refresh);
	const body: HTMLElement[] = [tiles, groupingField];
	if (!change) body.push(field(ctx, 'Chart title', title));
	body.push(preview);
	if (!change && t.ws.rows.size === 0)
		body.push(text(ctx, 'Select the data for the chart first, then insert it.'));
	refresh();
	return showDialog<ChartObject>(ctx, {
		name: 'insert-chart',
		heading: change ? 'Change Chart Type' : 'Insert Chart',
		wide: true,
		body,
		opened: () => tiles.querySelector<HTMLButtonElement>(`[data-gallery-item="${type}"]`)?.focus(),
		submit: () => {
			const chart = current();
			if (change)
				editChart(ctx, () => ({
					chartType: type,
					...(type === 'column' || type === 'bar' ? { grouping: grouping.value as Grouping } : {}),
				}));
			else {
				const { kind: _kind, ...model } = chart;
				ctx.selection.set({ drawing: t.session.addChart(t.sheet, model) });
			}
			return chart;
		},
	});
}
