import type { RibbonHomeIntent } from 'ooxml-ui/pptx';
import { chartSelectionHomeControls, dispatchChartSelectionRequest } from 'ooxml-ui/pptx';
import React, { useCallback, useMemo } from 'react';
import { useTranslation } from 'react-i18next';

import { useChartFormatContext } from '../chart-format-context';
import { WebHomeControls } from './WebHomeControls';

/**
 * Chart Format > Current Selection: the shared
 * `pptx-ui-ribbon-home-chart-selection` strip (element drop-down, Format
 * Selection, Reset to Match Style) over the viewer's chart format commands.
 */
export function ChartSelectionStrip(): React.ReactElement {
	const { t } = useTranslation();
	const commands = useChartFormatContext();
	const chart = commands?.chart;
	const chartData = chart?.type === 'chart' ? chart.chartData : undefined;
	const controls = useMemo(
		() =>
			chartSelectionHomeControls({
				chartData,
				targetKey: commands?.target,
				editable: Boolean(commands?.editable),
				translate: t,
			}),
		[chartData, commands?.target, commands?.editable, t],
	);
	const request = useCallback(
		(_id: string, _part?: string, intent?: RibbonHomeIntent) => {
			if (!commands || !intent) {
				return;
			}
			dispatchChartSelectionRequest(intent, {
				select: commands.setTarget,
				formatSelection: () => commands.openPane(),
				reset: commands.reset,
			});
		},
		[commands],
	);
	return <WebHomeControls family='chart-selection' controls={controls} onRequest={request} />;
}
