import type { ChartDataSheetChangeEvent, PptxUiChartDataSheetElement } from 'ooxml-ui/pptx';
import type { PptxChartData } from 'pptx-viewer-core';
import { useEffect, useLayoutEffect, useRef } from 'react';
import { useTranslation } from 'react-i18next';

import { CARD } from './chart-panel-constants';

export interface ChartDataSheetProps {
	chartData: PptxChartData;
	canEdit: boolean;
	/** File > Options > Advanced > "Properties follow chart data point for current workbook". */
	followDataPoint: boolean;
	/**
	 * Cell to highlight, driven by the on-canvas chart part selection: a `pointIndex` highlights
	 * one value cell, series-only highlights the series name cell.
	 */
	highlightCell?: { seriesIndex: number; pointIndex?: number } | null;
	/** Commits the edited chart data (one history entry per sheet change). */
	onChange: (chartData: PptxChartData) => void;
}

/**
 * The chart's "Edit Data" sheet: the shared `pptx-ui-chart-data-sheet` element (categories x
 * series, keyboard navigation, tab-separated paste, locale numbers). This component only feeds
 * it the chart and commits what it emits.
 */
export function ChartDataSheet({
	chartData,
	canEdit,
	followDataPoint,
	highlightCell,
	onChange,
}: ChartDataSheetProps) {
	const { t } = useTranslation();
	const ref = useRef<PptxUiChartDataSheetElement>(null);

	useLayoutEffect(() => {
		if (ref.current) {
			ref.current.state = {
				chartData,
				canEdit,
				followDataPoint,
				highlight: highlightCell ?? null,
				translate: (key, params) => t(key, params),
			};
		}
	}, [chartData, canEdit, followDataPoint, highlightCell, t]);

	useEffect(() => {
		const sheet = ref.current;
		if (!sheet) {
			return;
		}
		const commit = (event: Event) =>
			onChange((event as ChartDataSheetChangeEvent).detail.chartData);
		sheet.addEventListener('pptx-chart-data-change', commit);
		return () => sheet.removeEventListener('pptx-chart-data-change', commit);
	}, [onChange]);

	return (
		<div className={CARD}>
			<pptx-ui-chart-data-sheet ref={ref} />
		</div>
	);
}
