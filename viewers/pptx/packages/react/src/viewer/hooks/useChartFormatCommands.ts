/**
 * useChartFormatCommands: owns the chart Format pane state (current chart
 * element, open flag) and builds the {@link ChartFormatCommands} the viewer
 * provides. Edits go through `updateElementById`, which records history.
 */
import { resetChartFormatTarget } from 'ooxml-ui/pptx';
import type { PptxChartData, PptxElement } from 'pptx-viewer-core';
import { useCallback, useEffect, useMemo, useState } from 'react';

import type { ChartFormatCommands } from '../components/chart-format-context';

export interface UseChartFormatCommandsInput {
	selectedElement: PptxElement | null | undefined;
	editable: boolean;
	updateElementById: (elementId: string, updates: Partial<PptxElement>) => void;
}

export function useChartFormatCommands(input: UseChartFormatCommandsInput): ChartFormatCommands {
	const { selectedElement, editable, updateElementById } = input;
	const chart = selectedElement?.type === 'chart' ? selectedElement : null;
	const chartId = chart?.id ?? null;
	// The element is remembered per chart, so another selection starts on the
	// chart area, while a right-click that selects a chart can name its element first.
	const [picked, setPicked] = useState<{ chartId: string | null; target: string }>({
		chartId: null,
		target: 'chartArea',
	});
	const [paneOpen, setPaneOpen] = useState(false);
	const target = picked.chartId === chartId ? picked.target : 'chartArea';
	useEffect(() => {
		if (!chartId) {
			setPaneOpen(false);
		}
	}, [chartId]);
	const setTarget = useCallback(
		(next: string, forChartId?: string) =>
			setPicked({ chartId: forChartId ?? chartId, target: next }),
		[chartId],
	);
	const openPane = useCallback(
		(next?: string) => {
			if (next) {
				setTarget(next);
			}
			setPaneOpen(true);
		},
		[setTarget],
	);
	const closePane = useCallback(() => setPaneOpen(false), []);
	const applyChartData = useCallback(
		(chartData: PptxChartData) => {
			if (editable && chartId) {
				updateElementById(chartId, { chartData } as Partial<PptxElement>);
			}
		},
		[editable, chartId, updateElementById],
	);
	const chartData = chart?.type === 'chart' ? chart.chartData : undefined;
	const reset = useCallback(() => {
		if (chartData) {
			applyChartData(resetChartFormatTarget(chartData, target));
		}
	}, [chartData, target, applyChartData]);
	return useMemo(
		() => ({
			chart,
			editable,
			target,
			paneOpen: paneOpen && Boolean(chart),
			setTarget,
			openPane,
			closePane,
			applyChartData,
			reset,
		}),
		[chart, editable, target, paneOpen, setTarget, openPane, closePane, applyChartData, reset],
	);
}
