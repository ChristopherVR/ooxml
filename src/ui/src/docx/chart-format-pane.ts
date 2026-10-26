// Word's chart Format pane: the shared `office-ui-chart-format-pane` docked beside the page for the
// selected chart (Format Selection on the Chart Design tab, Format Chart Area... on the chart's
// context menu, like Word 365). The pane reads the chart's `ChartSpace` through
// `chartSpaceFormatSource`; each committed field's neutral intent runs through
// `applyChartEditIntent` and the docx chart edit command, one undo step per edit, and is saved by
// the existing chart save path. Fields no neutral intent writes stay hidden (the source does not
// offer them) or are ignored.
import type { EditorView } from 'prosemirror-view';
import { applyChartExEditIntent, applyChartEditIntent } from 'ooxml-core/chart';
import {
	diagramColorTheme,
	editDocxChartEx,
	editDocxChart,
	type ThemeCatalog,
} from 'ooxml-core/docx';
import { selectedChart } from 'ooxml-core/docx/ui';
import {
	defineChartFormatPane,
	type ChartFormatEditEvent,
	type OfficeUiChartFormatPane,
} from '../chart/format-pane';
import { chartExSpaceFormatSource } from '../chart/format-pane-chartex';
import { editSelectedChart } from './chart-commands';
import { focusView } from './focus-view';
import type { EditorLocale } from './localization';

export interface ChartFormatPaneHost {
	view(): EditorView | undefined;
	theme(): ThemeCatalog | undefined;
	locale(): EditorLocale;
	/** Whether the body takes edits (not read-only). */
	canEdit(): boolean;
}

export interface DocxChartFormatPane {
	element: OfficeUiChartFormatPane;
	readonly isOpen: boolean;
	/** Opens the pane for the selected chart, on `target` (`chartArea` by default). */
	open(target?: string): void;
	close(): void;
	/** Repaints from the selected chart; closes the pane when no chart is selected. */
	refresh(): void;
}

/** The docked Format pane of the selected chart. */
export function createChartFormatPane(host: ChartFormatPaneHost): DocxChartFormatPane {
	defineChartFormatPane();
	const element = document.createElement('office-ui-chart-format-pane');
	element.className = 'dve-chart-format-pane';
	element.hidden = true;
	let locale: EditorLocale | undefined;
	let theme: ThemeCatalog | undefined;
	/** The chart node the pane shows (a chart's model is parsed again from each node). */
	let shown: unknown;

	const close = () => {
		if (element.hidden) return;
		element.hidden = true;
		shown = undefined;
		focusView(host.view());
	};
	const refresh = () => {
		if (element.hidden) return;
		const view = host.view();
		const selected = view && selectedChart(view.state);
		const chartSpace = selected?.chart.chartSpace;
		if (!view || !selected || (!chartSpace && !selected.chart.chartExSpace)) {
			close();
			return;
		}
		// The pane's captions come from its own dictionary (`chart/format-pane-locales`).
		const nextLocale = host.locale();
		if (nextLocale !== locale) {
			locale = nextLocale;
			element.locale = nextLocale;
		}
		const nextTheme = host.theme();
		if (nextTheme !== theme) {
			theme = nextTheme;
			element.theme = diagramColorTheme(nextTheme);
		}
		// A new model (an edit, undo or another chart) replaces the source; the target is kept.
		if (selected.node !== shown) {
			shown = selected.node;
			if (selected.chart.chartExSpace)
				element.source = chartExSpaceFormatSource(
					selected.chart.chartExSpace,
					diagramColorTheme(nextTheme),
				);
			else element.chartSpace = chartSpace;
		}
		element.disabled = !selected.editable || !view.editable || !host.canEdit();
	};

	element.addEventListener('office-chart-format-edit', (event) => {
		const { intent } = (event as ChartFormatEditEvent).detail;
		const view = host.view();
		if (intent && view && host.canEdit())
			try {
				editSelectedChart(view, (chart) =>
					chart.chartExSpace
						? editDocxChartEx(chart, (current) => applyChartExEditIntent(current, intent))
						: editDocxChart(chart, (current) => applyChartEditIntent(current, intent)),
				);
			} catch {
				// An edit the chart cannot take (a family without that element) leaves it as it was.
			}
		shown = undefined;
		refresh();
	});
	element.addEventListener('office-chart-format-close', close);

	return {
		element,
		get isOpen() {
			return !element.hidden;
		},
		open(target = 'chartArea') {
			const view = host.view();
			if (!view || !selectedChart(view.state)) return;
			element.hidden = false;
			shown = undefined;
			element.target = target;
			refresh();
			element.shadowRoot?.querySelector<HTMLElement>('[data-chart-format-target]')?.focus();
		},
		close,
		refresh,
	};
}
