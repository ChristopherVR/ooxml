/**
 * Builds the descriptor of PowerPoint's chart Format task pane for one chart
 * element: the shared chart Format pane model (`src/ui/src/chart`) over
 * {@link pptxChartFormatSource}.
 *
 * @module render/chart-format/chart-format-pane-model
 */
import type { PptxChartData, PptxChartFormatTargetKey } from 'ooxml-core/pptx';

import { buildChartFormatModel, chartFormatSourceTabs } from '../../../chart/format-pane-model';
import type {
	ChartFormatPaneMode,
	ChartFormatPaneModel,
	ChartFormatTab,
	ChartFormatTabId,
} from '../../../chart/format-pane-types';
import { pptxChartFormatSource } from './chart-format-source';
import { resolveChartFormatTarget } from './chart-format-targets';

/** The tabs `targetKey` shows in `mode`, in PowerPoint's order. */
export function chartFormatTabs(
	data: PptxChartData,
	targetKey: PptxChartFormatTargetKey | undefined,
	mode: ChartFormatPaneMode = 'chart',
): ChartFormatTab[] {
	const { target } = resolveChartFormatTarget(data, targetKey);
	return chartFormatSourceTabs(pptxChartFormatSource(data), target, mode);
}

/**
 * The pane for `targetKey` on `data` with `tab` open, in Chart Options or
 * Text Options `mode`. A stale target falls back to the chart area, Text
 * Options on an element without text to Chart Options, and a tab the
 * target lacks to the mode's first tab.
 */
export function buildChartFormatPaneModel(
	data: PptxChartData,
	targetKey: PptxChartFormatTargetKey | undefined,
	tab: ChartFormatTabId = 'fillLine',
	mode: ChartFormatPaneMode = 'chart',
): ChartFormatPaneModel {
	return buildChartFormatModel(pptxChartFormatSource(data), targetKey, tab, mode);
}
