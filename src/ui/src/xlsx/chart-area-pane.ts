import { chartView, createRefEvaluator, type ChartFillPart } from 'ooxml-core/xlsx';
import { activeChart, editing, icon, type Command, type EditorContext } from 'ooxml-core/xlsx/ui';
import { el, field, select } from './dialogs/fields';
import { createSeriesFill } from './chart-series-fill';
import { backgroundFillBinding } from './chart-fill-binding';

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
		title.textContent = ctx.t(selected === 'chartArea' ? 'Format Chart Area' : 'Format Plot Area');
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
			...(['chartArea', 'plotArea'] as const).map((part) => {
				const option = el(ctx, 'option');
				option.value = part;
				option.textContent = ctx.t(part === 'chartArea' ? 'Chart Area' : 'Plot Area');
				return option;
			}),
		);
		refresh();
	};
	target.addEventListener('change', () => {
		if (!target.disabled && (target.value === 'chartArea' || target.value === 'plotArea'))
			selected = target.value;
		refresh();
	});
	ctx.dialogs.register('format-chart-area', async (_ctx, part) => {
		if (!ctx.commands.isEnabled('chart.format-area')) return;
		onOpen();
		selected = part === 'plotArea' ? 'plotArea' : 'chartArea';
		element.hidden = false;
		refresh();
		target.focus();
	});
	relocalize();
	return { element, refresh, relocalize, close };
}
