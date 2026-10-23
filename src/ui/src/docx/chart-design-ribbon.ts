// The Chart Design tab: Word's contextual tab for a selected chart. It appears while a chart is
// selected and offers the edits the shared chart core supports here: the legend, Edit Data and
// Change Chart Type. Chart styles, colours and Quick Layout are not offered (there is no shared
// chart style gallery yet). The shared chart Format pane is to mount from this tab once it exists:
// add a "Format" group whose action opens it for `selectedChart(state)`.
import type { EditorState } from 'prosemirror-state';
import { selectedChart } from 'ooxml-core/docx/ui';
import { group, menuSelect, tool } from './ribbon-parts';
import { isTabHidden, panelOf, selectTab, setTabHidden } from './ribbon-tab-api';

const LEGEND: Array<[string, string]> = [
	['none', 'None'],
	['r', 'Right'],
	['t', 'Top'],
	['l', 'Left'],
	['b', 'Bottom'],
];

export function buildChartDesignPanel(panel: HTMLElement): void {
	panel.append(
		group(
			'Chart Layouts',
			menuSelect('Legend', 'legend', LEGEND, (value) => ({ type: 'chart', key: 'legend', value }), {
				momentary: true,
			}),
		),
		group(
			'Data',
			tool(
				'Edit Data',
				'editData',
				{ type: 'formatDialog', kind: 'chartData' },
				{ large: true, caption: 'Edit Data' },
			),
		),
		group(
			'Type',
			tool(
				'Change Chart Type',
				'chartType',
				{ type: 'formatDialog', kind: 'changeChartType' },
				{ large: true, caption: 'Change Chart Type' },
			),
		),
	);
}

/** Shows the tab while a chart is selected (switching to it on entry) and fills its controls. */
export function syncChartDesign(toolbar: HTMLElement, state: EditorState, editable: boolean): void {
	const panel = panelOf(toolbar, 'chart-design');
	if (!panel) return;
	const chart = selectedChart(state);
	const entering = isTabHidden(panel) && !!chart;
	setTabHidden(toolbar, 'chart-design', 'contextual', !chart);
	if (entering) selectTab(toolbar, 'chart-design');
	const active = Boolean(chart?.editable && editable);
	for (const control of panel.querySelectorAll<HTMLButtonElement | HTMLSelectElement>(
		'button, select',
	))
		control.disabled = !active;
	if (!chart) return;
	const legend = panel.querySelector<HTMLSelectElement>(
		'[data-localearialabel="Legend"], [aria-label="Legend"]',
	);
	if (legend) legend.value = chart.chart.chartSpace?.legend?.position ?? 'none';
}
