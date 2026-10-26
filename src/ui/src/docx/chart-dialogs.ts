// The chart dialogs of the Word editor, on the shared chart elements: Insert Chart and Change Chart
// Type host `office-ui-chart-picker`, Edit Data hosts `office-ui-chart-data-sheet`. The chart
// model and its edits are the shared core's; every change is one step of the editor's undo
// history and repaints the chart through the shared painter.
import type { EditorView } from 'prosemirror-view';
import {
	applyChartExEditAction,
	setChartLegend,
	setChartType,
	type ChartDataSheetEdit,
} from 'ooxml-core/chart';
import {
	applyDocxChartDataSheetEdits,
	docxChartDataSheet,
	docxChartPaint,
	editDocxChart,
	editDocxChartEx,
	newChartImage,
	newDocxChart,
	type DocxChart,
	type ThemeCatalog,
} from 'ooxml-core/docx';
import { insertChartImage, selectedChart } from 'ooxml-core/docx/ui';
import {
	defineChartExOptions,
	type OfficeUiChartExOptions,
	type ChartExEditEvent,
} from '../chart/chartex-options';
import { defineChartDataSheet, type OfficeChartDataSheetLabels } from '../chart/data-sheet';
import {
	defineChartPicker,
	type OfficeChartPickerLabels,
	type OfficeChartPreview,
	type OfficeChartValue,
} from '../dialog/chart-picker';
import {
	DOCX_CHART_GROUPINGS,
	DOCX_CHART_TYPES,
	DOCX_GROUPED_CHART_TYPES,
	chartPickerChoice,
	chartTypeTarget,
	editSelectedChart,
	setSelectedChartType,
} from './chart-commands';
import { sheetLabels } from './chart-sheet-labels';
import { dialogButton } from './dialog-fields';
import {
	createDialogShell,
	dialogActions,
	isDialogOpen,
	localizeDialog,
	onDialogDismiss,
	setDialogOpen,
} from './dialog-shell';
import { focusView } from './focus-view';
import type { FormatDialog } from './font-dialog';
import { localizeElement, translate, type EditorLocale } from './localization';

type Picker = HTMLElement & {
	types: { id: string; label: string }[];
	groupings: { id: string; label: string }[];
	groupedTypes: string[];
	labels: Partial<OfficeChartPickerLabels>;
	drawPreview: OfficeChartPreview;
	value: OfficeChartValue;
	changeOnly: boolean;
};

type Sheet = HTMLElement & {
	data: ReturnType<typeof docxChartDataSheet>;
	labels: Partial<OfficeChartDataSheetLabels>;
	locale: string;
	disabled: boolean;
};

export interface ChartDialogsHost {
	view(): EditorView | undefined;
	theme(): ThemeCatalog | undefined;
}

/** The chart a preview draws: `chart` (or a new sample chart) as the picker choice. */
function previewChart(
	value: OfficeChartValue,
	tile: boolean,
	chart: DocxChart | undefined,
): DocxChart {
	const target = chartTypeTarget(value.type, value.grouping);
	const title = value.title.trim();
	let next = chart ?? newDocxChart(target, title && !tile ? { title } : {});
	if (chart)
		try {
			next = editDocxChart(chart, (space) => setChartType(space, target));
		} catch {
			// A family this chart cannot switch to previews as the chart itself.
		}
	if (!tile) return next;
	// A type tile is drawn bare: no title, no legend.
	return editDocxChart(next, (space) => {
		const { title: _title, ...rest } = space;
		return setChartLegend({ ...rest, autoTitleDeleted: true }, { visible: false });
	});
}

function pickerLabels(locale: EditorLocale): Partial<OfficeChartPickerLabels> {
	return {
		chartType: translate(locale, 'Chart type'),
		grouping: translate(locale, 'Grouping'),
		title: translate(locale, 'Chart title'),
		data: translate(locale, 'Chart data'),
		dataHint: translate(locale, 'chart.dataHint'),
		preview: translate(locale, 'Preview'),
	};
}

/** Insert Chart (`change` false) or Change Chart Type (`change` true). */
export function createChartTypeDialog(host: ChartDialogsHost, change: boolean): FormatDialog {
	const title = change ? 'Change Chart Type' : 'Insert Chart';
	const element = createDialogShell(title, 'dve-format-dialog dve-chart-dialog');
	defineChartExOptions();
	const extended = document.createElement('office-ui-chart-ex-options') as OfficeUiChartExOptions;
	const extendedClose = dialogButton('OK', true);
	defineChartPicker();
	const picker = document.createElement('office-ui-chart-picker') as Picker;
	picker.changeOnly = change;
	const message = document.createElement('p');
	message.setAttribute('role', 'status');
	const cancel = dialogButton('Cancel');
	const ok = dialogButton('OK', true);
	let locale: EditorLocale = 'en';
	let current: DocxChart | undefined;
	const hide = () => {
		setDialogOpen(element, false);
		element.replaceChildren();
		current = undefined;
		focusView(host.view());
	};
	extendedClose.addEventListener('click', hide);
	extended.addEventListener('office-chart-ex-edit', (event) => {
		const view = host.view();
		if (!view) return;
		try {
			editSelectedChart(view, (chart) =>
				editDocxChartEx(chart, (space) =>
					applyChartExEditAction(space, (event as ChartExEditEvent).detail.action),
				),
			);
			message.textContent = '';
		} catch {
			message.textContent = translate(locale, 'chart.readOnly');
		}
		extended.space = selectedChart(view.state)?.chart.chartExSpace;
	});
	cancel.addEventListener('click', hide);
	onDialogDismiss(element, hide);
	ok.addEventListener('click', () => {
		const view = host.view();
		const { type, grouping, title: text } = picker.value;
		if (view?.editable) {
			if (change) setSelectedChartType(view, type, grouping);
			else {
				const chart = newDocxChart(
					chartTypeTarget(type, grouping),
					text.trim() ? { title: text.trim() } : {},
				);
				const tr = insertChartImage(view.state, newChartImage(chart));
				if (tr) view.dispatch(tr);
			}
		}
		hide();
	});
	picker.drawPreview = (value, size, tile) => {
		try {
			const chart = previewChart(value, tile, current);
			return docxChartPaint(chart, size.width, size.height, host.theme()).svg ?? '';
		} catch {
			return '';
		}
	};
	return {
		element,
		close: hide,
		setLocale(next) {
			locale = next;
		},
		get isOpen() {
			return isDialogOpen(element);
		},
		open() {
			const view = host.view();
			if (!view?.editable) return;
			const selected = change ? selectedChart(view.state) : undefined;
			if (change && selected?.editable && selected.chart.chartExSpace) {
				extended.space = selected.chart.chartExSpace;
				extended.locale = locale;
				message.textContent = '';
				element.replaceChildren(extended, message, dialogActions(extendedClose));
				localizeDialog(element, locale);
				setDialogOpen(element, true);
				extendedClose.focus();
				return;
			}
			if (change && !selected?.editable) {
				message.textContent = translate(locale, 'chart.readOnly');
				element.replaceChildren(message, dialogActions(cancel));
			} else {
				current = selected?.chart;
				picker.types = DOCX_CHART_TYPES.map(([id, label]) => ({
					id,
					label: translate(locale, label),
				}));
				picker.groupings = DOCX_CHART_GROUPINGS.map(([id, label]) => ({
					id,
					label: translate(locale, label),
				}));
				picker.groupedTypes = [...DOCX_GROUPED_CHART_TYPES];
				picker.labels = pickerLabels(locale);
				picker.value = current
					? { ...chartPickerChoice(current), title: current.title ?? '' }
					: { type: 'column', grouping: 'clustered', title: translate(locale, 'Chart title') };
				element.replaceChildren(picker, dialogActions(cancel, ok));
			}
			localizeDialog(element, locale);
			setDialogOpen(element, true);
			(change && !current ? cancel : ok).focus();
		},
	};
}

/** Edit Data: the selected chart's data sheet; each edit is applied at once as one undo step. */
export function createChartDataDialog(host: ChartDialogsHost): FormatDialog {
	const element = createDialogShell('Edit Data', 'dve-format-dialog dve-chart-data-dialog');
	defineChartDataSheet();
	defineChartExOptions();
	const extended = document.createElement('office-ui-chart-ex-options') as OfficeUiChartExOptions;
	const sheet = document.createElement('office-ui-chart-data-sheet') as Sheet;
	const message = document.createElement('p');
	message.setAttribute('role', 'status');
	const close = dialogButton('OK', true);
	let locale: EditorLocale = 'en';
	const hide = () => {
		setDialogOpen(element, false);
		element.replaceChildren();
		focusView(host.view());
	};
	close.addEventListener('click', hide);
	onDialogDismiss(element, hide);
	sheet.addEventListener('office-chart-data-change', (event) => {
		const view = host.view();
		const { edits } = (event as CustomEvent<{ edits: ChartDataSheetEdit[] }>).detail;
		if (!view) return;
		try {
			editSelectedChart(view, (chart) => applyDocxChartDataSheetEdits(chart, edits));
			message.textContent = '';
		} catch {
			message.textContent = translate(locale, 'chart.readOnly');
		}
		extended.space = selectedChart(view.state)?.chart.chartExSpace;
		const chart = selectedChart(view.state)?.chart;
		const data = docxChartDataSheet(chart);
		if (data) sheet.data = data;
	});
	extended.addEventListener('office-chart-ex-edit', (event) => {
		const view = host.view();
		if (!view) return;
		try {
			editSelectedChart(view, (chart) =>
				editDocxChartEx(chart, (space) =>
					applyChartExEditAction(space, (event as ChartExEditEvent).detail.action),
				),
			);
			message.textContent = '';
		} catch {
			message.textContent = translate(locale, 'chart.readOnly');
		}
		extended.space = selectedChart(view.state)?.chart.chartExSpace;
		const data = docxChartDataSheet(selectedChart(view.state)?.chart);
		if (data) sheet.data = data;
	});
	return {
		element,
		close: hide,
		setLocale(next) {
			locale = next;
		},
		get isOpen() {
			return isDialogOpen(element);
		},
		open() {
			const view = host.view();
			if (!view?.editable) return;
			const data = docxChartDataSheet(selectedChart(view.state)?.chart);
			if (!data) {
				message.textContent = translate(locale, 'chart.readOnly');
				element.replaceChildren(message, dialogActions(close));
			} else {
				sheet.labels = sheetLabels(locale);
				sheet.locale = locale;
				sheet.data = data;
				const selected = selectedChart(view.state)?.chart;
				extended.space = selected?.chartExSpace;
				extended.locale = locale;
				sheet.disabled =
					!!selected?.chartExSpace?.chartData.externalData && !selected.workbookPartName;
				message.textContent = sheet.disabled ? translate(locale, 'chart.readOnly') : '';
				element.replaceChildren(
					...(selected?.chartExSpace ? [extended] : []),
					sheet,
					message,
					dialogActions(close),
				);
			}
			localizeDialog(element, locale);
			localizeElement(message, locale);
			setDialogOpen(element, true);
			close.focus();
		},
	};
}
