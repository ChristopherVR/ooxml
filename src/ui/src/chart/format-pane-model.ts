/**
 * Builds the descriptor of a chart Format task pane for one chart element
 * from a {@link ChartFormatSource}. Moved from PowerPoint's
 * `render/chart-format/chart-format-pane-model.ts`, which now passes its own
 * source.
 *
 * @module chart/format-pane-model
 */
import type { ChartFormatTarget } from 'ooxml-core/chart';

import { t } from './format-pane-common';
import { effectsSections } from './format-pane-effects';
import { fillLineSections } from './format-pane-fill-line';
import { optionsSections } from './format-pane-options';
import { CHART_AREA_TARGET, chartFormatPaneTitle } from './format-pane-targets';
import { chartTextTabs, textOptionsSections } from './format-pane-text';
import type {
	ChartFormatPaneMode,
	ChartFormatPaneModel,
	ChartFormatSection,
	ChartFormatTab,
	ChartFormatTabId,
	ChartFormatTargetOption,
} from './format-pane-types';
import type { ChartFormatElementView, ChartFormatSource } from './format-pane-view';

/** `key` when it names an element of `source`, else its first target (the chart area). */
export function resolveChartFormatSourceTarget(
	source: ChartFormatSource | undefined,
	key: string | undefined,
): ChartFormatTargetOption {
	const match = key ? source?.targets.find((option) => option.key === key) : undefined;
	return (
		match ?? (key ? source?.resolve?.(key) : undefined) ?? source?.targets[0] ?? CHART_AREA_TARGET
	);
}

const hasText = (view: ChartFormatElementView) => view.support.text && view.text !== undefined;
const hasEffects = (view: ChartFormatElementView) =>
	view.support.effects !== false && view.effects !== undefined;

function optionsLabel(target: ChartFormatTarget) {
	switch (target.kind) {
		case 'series':
			return t('tab.seriesOptions', 'Series Options');
		case 'axis':
			return t('tab.axisOptions', 'Axis Options');
		case 'legend':
			return t('tab.legendOptions', 'Legend Options');
		default:
			return t('tab.labelOptions', 'Label Options');
	}
}

/** The tabs `target` shows in `mode`, in Office's order. */
export function chartFormatSourceTabs(
	source: ChartFormatSource,
	target: ChartFormatTarget,
	mode: ChartFormatPaneMode = 'chart',
): ChartFormatTab[] {
	const view = source.element(target);
	if (mode === 'text' && hasText(view)) return chartTextTabs();
	const tabs: ChartFormatTab[] = [{ id: 'fillLine', label: t('tab.fillLine', 'Fill & Line') }];
	if (hasEffects(view)) tabs.push({ id: 'effects', label: t('tab.effects', 'Effects') });
	if (optionsSections(source, target)) tabs.push({ id: 'options', label: optionsLabel(target) });
	return tabs;
}

function tabSections(
	source: ChartFormatSource,
	target: ChartFormatTarget,
	view: ChartFormatElementView,
	tab: ChartFormatTabId,
): ChartFormatSection[] {
	switch (tab) {
		case 'textFill':
		case 'textEffects':
		case 'textBox':
			return view.text ? textOptionsSections(view.text, tab) : [];
		case 'effects':
			return hasEffects(view) ? effectsSections(view) : [];
		case 'options':
			return optionsSections(source, target) ?? [];
		default:
			return fillLineSections(view, target, source);
	}
}

/**
 * The pane for `targetKey` with `tab` open, in Chart Options or Text Options
 * `mode`. A stale target falls back to the source's first target, Text Options
 * on an element without text to Chart Options, and a tab the target lacks to
 * the mode's first tab. A `flat` source gets every chart section on one page.
 */
export function buildChartFormatModel(
	source: ChartFormatSource,
	targetKey: string | undefined,
	tab: ChartFormatTabId = 'fillLine',
	mode: ChartFormatPaneMode = 'chart',
): ChartFormatPaneModel {
	const resolved = resolveChartFormatSourceTarget(source, targetKey);
	const view = source.element(resolved.target);
	const activeMode: ChartFormatPaneMode = mode === 'text' && hasText(view) ? 'text' : 'chart';
	const tabs = chartFormatSourceTabs(source, resolved.target, activeMode);
	const activeTab = tabs.some((candidate) => candidate.id === tab) ? tab : tabs[0]!.id;
	const flat = Boolean(source.flat) && activeMode === 'chart';
	const sections = flat
		? tabs.flatMap((entry) => tabSections(source, resolved.target, view, entry.id))
		: tabSections(source, resolved.target, view, activeTab);
	return {
		title: chartFormatPaneTitle(resolved.target),
		targetKey: resolved.key,
		target: resolved.target,
		targets: [...source.targets],
		picker: source.picker ?? t('chartOptions', 'Chart Options'),
		tabs: flat ? [] : tabs,
		activeTab: flat ? 'fillLine' : activeTab,
		sections,
		mode: activeMode,
		textOptions: hasText(view),
	};
}
