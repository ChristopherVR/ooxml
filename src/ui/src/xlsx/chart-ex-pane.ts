// A preserved ChartEx drawing opens shared options, data and formatting controls on double-click.
import {
	applyChartExEditAction,
	applyChartExEditIntent,
	applyChartExSpaceDataSheetEdits,
	type ChartExSpace,
} from 'ooxml-core/chart';
import { xlsxChartExDataSheet } from 'ooxml-core/xlsx';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import {
	defineChartExOptions,
	type OfficeUiChartExOptions,
	type ChartExEditEvent,
} from '../chart/chartex-options';
import {
	defineChartDataSheet,
	type OfficeUiChartDataSheet,
	type OfficeChartDataChangeEvent,
} from '../chart/data-sheet';
import {
	defineChartFormatPane,
	type OfficeUiChartFormatPane,
	type ChartFormatEditEvent,
	type ChartFormatTargetEvent,
} from '../chart/format-pane';
import { chartExSpaceFormatSource } from '../chart/format-pane-chartex';
import { chartDataSheetLabels } from '../chart/data-sheet-locales';
import { workbookChartTheme } from './chart-format-source';

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
export function createChartExPane(ctx: EditorContext) {
	const doc = ctx.host.ownerDocument;
	defineChartExOptions(doc.defaultView?.customElements);
	defineChartDataSheet(doc.defaultView?.customElements);
	defineChartFormatPane(doc.defaultView?.customElements);
	const element = doc.createElement('aside');
	element.className = 'xve-chart-series-pane';
	element.hidden = true;
	const options = doc.createElement('office-ui-chart-ex-options') as OfficeUiChartExOptions;
	const sheet = doc.createElement('office-ui-chart-data-sheet') as OfficeUiChartDataSheet;
	const format = doc.createElement('office-ui-chart-format-pane') as OfficeUiChartFormatPane;
	const closeButton = doc.createElement('button');
	closeButton.type = 'button';
	const error = doc.createElement('p');
	error.setAttribute('role', 'status');
	const close = () => {
		element.hidden = true;
		ctx.grid()?.focus();
	};
	closeButton.addEventListener('click', close);
	element.append(closeButton, options, sheet, error, format);
	const refresh = () => {
		if (element.hidden) return;
		const found = activeChartEx(ctx),
			book = ctx.workbook();
		if (!found || !book) return close();
		options.space = found.drawing.chartExSpace;
		options.locale = ctx.locale();
		options.disabled = ctx.readOnly() || !ctx.commands.isEnabled('chart.format-area');
		format.source = chartExSpaceFormatSource(found.drawing.chartExSpace!, workbookChartTheme(book));
		format.locale = ctx.locale();
		format.disabled = options.disabled;
		closeButton.textContent = ctx.t('Close');
		element.setAttribute('aria-label', ctx.t('Chart'));
		try {
			sheet.data = xlsxChartExDataSheet(book, ctx.activeSheet(), found.index);
			sheet.disabled = options.disabled;
			sheet.labels = chartDataSheetLabels(ctx.locale());
			sheet.locale = ctx.locale();
			sheet.hidden = false;
		} catch {
			sheet.hidden = true;
		}
	};
	const apply = (operation: (space: ChartExSpace) => ChartExSpace) => {
		const found = activeChartEx(ctx);
		if (!found || ctx.readOnly() || !ctx.commands.isEnabled('chart.format-area')) return;
		try {
			ctx.session()?.editChartEx(ctx.activeSheet(), found.index, operation);
			error.textContent = '';
		} catch {
			error.textContent = ctx.t('This chart element cannot take that setting.');
		}
		refresh();
	};
	options.addEventListener('office-chart-ex-edit', (event) =>
		apply((space) => applyChartExEditAction(space, (event as ChartExEditEvent).detail.action)),
	);
	sheet.addEventListener('office-chart-data-change', (event) =>
		apply((space) =>
			applyChartExSpaceDataSheetEdits(space, (event as OfficeChartDataChangeEvent).detail.edits),
		),
	);
	format.addEventListener('office-chart-format-edit', (event) => {
		const intent = (event as ChartFormatEditEvent).detail.intent;
		if (intent) apply((space) => applyChartExEditIntent(space, intent));
	});
	format.addEventListener('office-chart-format-target', (event) => {
		format.target = (event as ChartFormatTargetEvent).detail.target;
	});
	format.addEventListener('office-chart-format-close', close);
	ctx.dialogs.register('chart-ex', async () => {
		if (!activeChartEx(ctx) || ctx.readOnly()) return;
		element.hidden = false;
		error.textContent = '';
		refresh();
	});
	return { element, refresh, close, relocalize: refresh };
}
