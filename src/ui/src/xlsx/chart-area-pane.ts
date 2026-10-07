import {
	chartView,
	createRefEvaluator,
	CHART_FILL_PARTS,
	type ChartFillPart,
} from 'ooxml-core/xlsx';
import { activeChart, editing, icon, type Command, type EditorContext } from 'ooxml-core/xlsx/ui';
import { el, field, select } from './dialogs/fields';
import { createSeriesFill } from './chart-series-fill';
import { backgroundFillBinding } from './chart-fill-binding';

const names: Record<ChartFillPart, [string, string]> = {
	chartArea: ['Chart Area', 'Format Chart Area'],
	plotArea: ['Plot Area', 'Format Plot Area'],
	title: ['Chart Title', 'Format Chart Title'],
	legend: ['Legend', 'Format Legend'],
};

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

/** Background targets share the series fill controls and all their native Office galleries. */
export function createChartAreaPane(ctx: EditorContext, onOpen: () => void) {
	const element = el(ctx, 'aside', 'xve-chart-series-pane');
	element.hidden = true;
	element.setAttribute('role', 'complementary');
	const header = el(ctx, 'div', 'xve-chart-series-header');
	const title = el(ctx, 'h2');
	const closeButton = el(ctx, 'button', 'xve-icon-button');
	closeButton.type = 'button';
	closeButton.textContent = '×';
	const body = el(ctx, 'div', 'xve-chart-series-body');
	let selected: ChartFillPart = 'chartArea';
	const target = select(ctx, []);
	const targetField = field(ctx, 'Chart element', target);
	const fills = createSeriesFill(
		ctx,
		() => 0,
		backgroundFillBinding(() => selected),
	);
	const close = () => {
		fills.refresh(undefined, undefined);
		element.hidden = true;
		ctx.grid()?.focus();
	};
	closeButton.addEventListener('click', close);
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') {
			event.preventDefault();
			event.stopPropagation();
			close();
		}
	});
	header.append(title, closeButton);
	body.append(targetField, fills.element);
	element.append(header, body);
	const refresh = () => {
		if (element.hidden) return;
		const found = activeChart(ctx),
			book = ctx.workbook();
		if (!found || !book) {
			close();
			return;
		}
		const available = (part: ChartFillPart) =>
			part === 'title' ? !!found.chart.title : part === 'legend' ? found.chart.showLegend : true;
		for (const option of target.options)
			option.disabled = !available(option.value as ChartFillPart);
		if (!available(selected)) selected = 'chartArea';
		title.textContent = ctx.t(names[selected][1]);
		element.setAttribute('aria-label', title.textContent);
		target.value = selected;
		target.disabled = !ctx.commands.isEnabled('chart.format-area');
		fills.refresh(
			found.chart,
			chartView(
				book,
				ctx.activeSheet(),
				found.chart,
				createRefEvaluator(book, ctx.session()?.calc, { sheet: ctx.activeSheet() }),
			),
		);
	};
	const relocalize = () => {
		closeButton.setAttribute('aria-label', ctx.t('Close'));
		targetField.querySelector('span')!.textContent = ctx.t('Chart element');
		target.setAttribute('aria-label', ctx.t('Chart element'));
		target.replaceChildren(
			...CHART_FILL_PARTS.map((part) => {
				const option = el(ctx, 'option');
				option.value = part;
				option.textContent = ctx.t(names[part][0]);
				return option;
			}),
		);
		refresh();
	};
	target.addEventListener('change', () => {
		if (!target.disabled && CHART_FILL_PARTS.includes(target.value as ChartFillPart))
			selected = target.value as ChartFillPart;
		refresh();
	});
	ctx.dialogs.register('format-chart-area', async (_ctx, part) => {
		if (!ctx.commands.isEnabled('chart.format-area')) return;
		onOpen();
		selected = CHART_FILL_PARTS.includes(part as ChartFillPart)
			? (part as ChartFillPart)
			: 'chartArea';
		element.hidden = false;
		refresh();
		target.focus();
	});
	relocalize();
	return { element, refresh, relocalize, close };
}
