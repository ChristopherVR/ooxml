<script setup lang="ts">
import type { ChartDataSheetChangeEvent, ChartDataSheetViewState } from 'ooxml-ui/pptx';
import type { PptxChartData } from 'pptx-viewer-core';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

/**
 * ChartDataSheet: the chart's "Edit Data" sheet, the shared `pptx-ui-chart-data-sheet` element
 * (categories x series, keyboard navigation, tab-separated paste, locale numbers). This
 * component only feeds it the chart and emits the edited chart data, which the panel commits
 * through its `update` contract.
 */
const props = defineProps<{
	chartData: PptxChartData;
	canEdit: boolean;
	followDataPoint: boolean;
	highlightCell?: { seriesIndex: number; pointIndex?: number } | null;
}>();

const emit = defineEmits<{
	change: [chartData: PptxChartData];
}>();

const { t } = useI18n();

const state = computed<ChartDataSheetViewState>(() => ({
	chartData: props.chartData,
	canEdit: props.canEdit,
	followDataPoint: props.followDataPoint,
	highlight: props.highlightCell ?? null,
	translate: (key, params) => (params ? t(key, { ...params }) : t(key)),
}));

function onChange(event: Event): void {
	emit('change', (event as ChartDataSheetChangeEvent).detail.chartData);
}
</script>

<template>
	<div class="pptx-vue-chart-card rounded border border-border bg-card p-2">
		<pptx-ui-chart-data-sheet :state.prop="state" @pptx-chart-data-change="onChange" />
	</div>
</template>
