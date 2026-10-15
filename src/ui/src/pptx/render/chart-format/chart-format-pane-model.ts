/**
 * Builds the descriptor of a chart Format task pane for one chart element.
 *
 * @module render/chart-format/chart-format-pane-model
 */
import type { PptxChartData, PptxChartFormatTargetKey } from 'ooxml-core/pptx';
import { readChartElementFormat } from 'ooxml-core/pptx';

import { effectsSections } from './chart-format-pane-effects';
import { fillLineSections, t } from './chart-format-pane-fill-line';
import { hasOptionsTab, optionsSections } from './chart-format-pane-options';
import type {
	ChartFormatPaneModel,
	ChartFormatTab,
	ChartFormatTabId,
} from './chart-format-pane-types';
import {
	chartFormatPaneTitle,
	listChartFormatTargets,
	resolveChartFormatTarget,
} from './chart-format-targets';

/** The tabs `targetKey` shows, in PowerPoint's order. */
export function chartFormatTabs(
	data: PptxChartData,
	targetKey: PptxChartFormatTargetKey | undefined,
): ChartFormatTab[] {
	const { target } = resolveChartFormatTarget(data, targetKey);
	const tabs: ChartFormatTab[] = [
		{ id: 'fillLine', label: t('tab.fillLine', 'Fill & Line') },
		{ id: 'effects', label: t('tab.effects', 'Effects') },
	];
	if (hasOptionsTab(target)) {
		const label =
			target.kind === 'series'
				? t('tab.seriesOptions', 'Series Options')
				: target.kind === 'axis'
					? t('tab.axisOptions', 'Axis Options')
					: target.kind === 'legend'
						? t('tab.legendOptions', 'Legend Options')
						: t('tab.labelOptions', 'Label Options');
		tabs.push({ id: 'options', label });
	}
	return tabs;
}

/**
 * The pane for `targetKey` on `data` with `tab` open. A stale target falls
 * back to the chart area and a tab the target lacks to Fill & Line.
 */
export function buildChartFormatPaneModel(
	data: PptxChartData,
	targetKey: PptxChartFormatTargetKey | undefined,
	tab: ChartFormatTabId = 'fillLine',
): ChartFormatPaneModel {
	const resolved = resolveChartFormatTarget(data, targetKey);
	const tabs = chartFormatTabs(data, resolved.key);
	const activeTab = tabs.some((candidate) => candidate.id === tab) ? tab : 'fillLine';
	const view = readChartElementFormat(data, resolved.target);
	const sections =
		activeTab === 'fillLine'
			? fillLineSections(view, resolved.target)
			: activeTab === 'effects'
				? effectsSections(view)
				: optionsSections(data, resolved.target);
	return {
		title: chartFormatPaneTitle(resolved.target),
		targetKey: resolved.key,
		targets: listChartFormatTargets(data),
		tabs,
		activeTab,
		sections,
	};
}
