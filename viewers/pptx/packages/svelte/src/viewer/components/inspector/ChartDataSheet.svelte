<script lang="ts">
	/**
	 * ChartDataSheet: the chart's "Edit Data" sheet, the shared `pptx-ui-chart-data-sheet`
	 * element (categories x series, keyboard navigation, tab-separated paste, locale numbers).
	 * This component only feeds it the chart and hands the edited chart data to `onreplace`,
	 * which commits it as one element patch.
	 */
	import type { ChartDataSheetChangeEvent, ChartDataSheetViewState } from 'ooxml-ui/pptx';
	import type { PptxChartData } from 'pptx-viewer-core';

	import { useTranslator } from '../../../i18n/context';
	import { useViewerOptions } from '../../state/viewer-options-context';

	const {
		data,
		canEdit,
		onreplace,
	}: {
		data: PptxChartData;
		canEdit: boolean;
		onreplace: (next: PptxChartData) => void;
	} = $props();
	const t = useTranslator();
	const optionsState = useViewerOptions();

	const sheetState = $derived<ChartDataSheetViewState>({
		chartData: data,
		canEdit,
		// File > Options > Advanced > "Properties follow chart data point for current workbook".
		followDataPoint: optionsState.options.advanced.chartPropertiesFollowDataPoint,
		translate: t,
	});

	function onchange(event: Event): void {
		onreplace((event as ChartDataSheetChangeEvent).detail.chartData);
	}
</script>

<pptx-ui-chart-data-sheet
	class="pptx-svelte-chart-data-sheet"
	state={sheetState}
	onpptx-chart-data-change={onchange}
></pptx-ui-chart-data-sheet>

<style>
	.pptx-svelte-chart-data-sheet {
		margin-top: 8px;
		padding-top: 7px;
		border-top: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
	}
</style>
