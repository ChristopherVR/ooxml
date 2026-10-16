import type {
	ChartFormatChangeEvent,
	ChartFormatTargetEvent,
	PptxUiChartFormatPaneElement,
} from 'ooxml-ui/pptx';
import { useEffect, useRef } from 'react';
import { useTranslation } from 'react-i18next';

import { useChartFormatContext } from '../chart-format-context';

/**
 * React adapter for the shared chart Format task pane: it feeds the selected
 * chart, the current element and the translator in, and routes the pane's
 * edits, retargets and close back through the viewer's chart format commands.
 */
export function ChartFormatPanel() {
	const { t } = useTranslation();
	const commands = useChartFormatContext();
	const ref = useRef<PptxUiChartFormatPaneElement>(null);
	const chart = commands?.chart;
	const chartData = chart?.type === 'chart' ? chart.chartData : undefined;
	useEffect(() => {
		const node = ref.current;
		if (node && commands) {
			node.translateLabel = t;
			node.chartData = chartData;
			node.target = commands.target;
			node.disabled = !commands.editable;
		}
	}, [chartData, commands, t]);
	useEffect(() => {
		const node = ref.current;
		if (!node || !commands) {
			return;
		}
		const change = (event: Event) =>
			commands.applyChartData((event as ChartFormatChangeEvent).detail.chartData);
		const retarget = (event: Event) =>
			commands.setTarget((event as ChartFormatTargetEvent).detail.target);
		node.addEventListener('chart-format-change', change);
		node.addEventListener('chart-format-target', retarget);
		node.addEventListener('chart-format-close', commands.closePane);
		return () => {
			node.removeEventListener('chart-format-change', change);
			node.removeEventListener('chart-format-target', retarget);
			node.removeEventListener('chart-format-close', commands.closePane);
		};
	}, [commands]);
	if (!commands?.paneOpen || !chartData) {
		return null;
	}
	return <pptx-ui-chart-format-pane ref={ref} data-testid='chart-format-pane' />;
}
