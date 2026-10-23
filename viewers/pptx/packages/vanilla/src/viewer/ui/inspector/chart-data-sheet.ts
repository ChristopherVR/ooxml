import { registerPptxWebControls } from 'ooxml-ui/pptx';
import type { ChartDataSheetChangeEvent, PptxUiChartDataSheetElement } from 'ooxml-ui/pptx';
import type { PptxChartData } from 'pptx-viewer-core';

import type { Translator } from '../../i18n';

/** The on-canvas chart part selection: a `pointIndex` marks a value cell, none the series name. */
export interface ChartHighlightCell {
	seriesIndex: number;
	pointIndex?: number;
}

/**
 * The chart's "Edit Data" sheet: the shared `pptx-ui-chart-data-sheet` element (categories x
 * series, keyboard navigation, tab-separated paste, locale numbers), as every binding shows it.
 * This view only feeds it the chart and commits the edited chart data through `onChange`.
 */
export function createChartDataSheet(
	doc: Document,
	t: Translator,
	onChange: (data: PptxChartData) => void,
	/**
	 * File > Options > Advanced > "Properties follow chart data point for current workbook",
	 * read fresh on every update. Defaults to PowerPoint's own default (`true`).
	 */
	getFollowDataPoint: () => boolean = () => true,
) {
	registerPptxWebControls();
	const el = doc.createElement('pptx-ui-chart-data-sheet') as PptxUiChartDataSheetElement;
	el.classList.add('pptxv-chart-sheet');
	el.addEventListener('pptx-chart-data-change', (event) =>
		onChange((event as ChartDataSheetChangeEvent).detail.chartData),
	);
	let editable = true;
	let data: PptxChartData | undefined;
	let highlight: ChartHighlightCell | null = null;
	const render = (): void => {
		el.state = {
			chartData: data,
			canEdit: editable,
			followDataPoint: getFollowDataPoint(),
			highlight,
			translate: t,
		};
	};

	return {
		el,
		update(next: PptxChartData | undefined, highlightCell: ChartHighlightCell | null = null) {
			data = next;
			highlight = highlightCell;
			render();
		},
		setEditable(next: boolean) {
			editable = next;
			render();
		},
	};
}
