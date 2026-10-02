// Insert Chart: a chart type gallery, grouping, title and a live preview drawn by the core's
// chartView + renderChartSvg. OK adds a ChartObject over the selection (or, from Chart Design,
// changes the active chart's type).
import {
	type ChartObject,
	type ChartType,
	chartView,
	renderChartSvg,
} from '@christophervr/xlsx-core';
import { activeChart, editChart } from '../commands/contextual.js';
import { CHART_TYPES } from '../commands/insert.js';
import { regionOf, target } from '../commands/util.js';
import type { EditorContext } from '../context.js';
import { el, field, select, text, textInput } from './fields.js';
import { showDialog } from './frame.js';
import { buildChart, evaluateRef } from './insert-chart-model.js';

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

/** Parses the core's SVG string into nodes (no innerHTML). */
function svgNode(ctx: EditorContext, markup: string): Node | undefined {
	const Parser = ctx.host.ownerDocument.defaultView?.DOMParser;
	if (!Parser) return undefined;
	const doc = new Parser().parseFromString(markup, 'image/svg+xml');
	if (doc.querySelector('parsererror')) return undefined;
	return ctx.host.ownerDocument.importNode(doc.documentElement, true);
}

export function openInsertChart(
	ctx: EditorContext,
	props: InsertChartProps = {},
): Promise<ChartObject | undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const change = props.change === true ? activeChart(ctx) : undefined;
	let type: ChartType = props.type ?? change?.chart.chartType ?? 'column';
	const range = regionOf(t);
	const tiles = el(ctx, 'div', 'xve-tiles');
	tiles.setAttribute('role', 'group');
	tiles.setAttribute('aria-label', ctx.t('Chart type'));
	const grouping = select(ctx, GROUPINGS, change?.chart.grouping ?? 'clustered');
	const groupingField = field(ctx, 'Grouping', grouping);
	const title = textInput(ctx, change?.chart.title ?? '');
	const preview = el(ctx, 'div', 'xve-preview');
	preview.setAttribute('aria-label', ctx.t('Preview'));
	preview.setAttribute('role', 'img');
	const buttons = new Map<ChartType, HTMLButtonElement>();
	const current = (): ChartObject => {
		if (change) return { ...structuredClone(change.chart), chartType: type };
		return buildChart(t.workbook, t.sheet, range, type, {
			grouping: grouping.value as Grouping,
			...(title.value.trim() ? { title: title.value.trim() } : {}),
		});
	};
	const refresh = (): void => {
		for (const [key, button] of buttons) button.setAttribute('aria-pressed', String(key === type));
		groupingField.hidden = !!change || !(type === 'column' || type === 'bar');
		const model = chartView(t.workbook, t.sheet, current(), (ref) =>
			evaluateRef(t.workbook, t.sheet, ref),
		);
		const node = svgNode(ctx, renderChartSvg(model, 360, 220));
		preview.replaceChildren(...(node ? [node] : []));
	};
	for (const [key, label] of CHART_TYPES) {
		const button = el(ctx, 'button', 'xve-tile');
		button.type = 'button';
		button.textContent = ctx.t(label);
		button.dataset.type = key;
		button.addEventListener('click', () => {
			type = key;
			refresh();
		});
		buttons.set(key, button);
		tiles.append(button);
	}
	tiles.addEventListener('keydown', (event) => {
		const list = [...buttons.values()];
		const i = list.indexOf(event.target as HTMLButtonElement);
		const delta =
			event.key === 'ArrowRight' || event.key === 'ArrowDown'
				? 1
				: event.key === 'ArrowLeft' || event.key === 'ArrowUp'
					? -1
					: 0;
		if (i < 0 || !delta) return;
		event.preventDefault();
		const next = list[(i + delta + list.length) % list.length];
		next?.focus();
		next?.click();
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
		opened: () => buttons.get(type)?.focus(),
		submit: () => {
			const chart = current();
			if (change) editChart(ctx, () => ({ chartType: type }));
			else {
				const { kind: _kind, ...model } = chart;
				ctx.selection.set({ drawing: t.session.addChart(t.sheet, model) });
			}
			return chart;
		},
	});
}
