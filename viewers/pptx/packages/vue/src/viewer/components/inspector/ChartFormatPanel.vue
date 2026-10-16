<script setup lang="ts">
/**
 * ChartFormatPanel: hosts the shared `pptx-ui-chart-format-pane` (Format Chart
 * Area / Data Series / Axis / ...). The element owns the view and the field
 * edits; this maps the injected chart-format state onto it and routes a
 * committed chart model through the history-tracked update.
 */
import type { ChartFormatChangeEvent, ChartFormatTargetEvent } from 'ooxml-ui/pptx';
import { computed, toRaw } from 'vue';
import { useI18n } from 'vue-i18n';

import { injectChartFormat } from '../../composables/chart-format-pane';

const controller = injectChartFormat();
const { t, locale } = useI18n();
const visible = computed(() => Boolean(controller?.open.value && controller.chart.value));
// A new translator identity per locale makes the element re-read its labels.
const translate = computed(() => {
	void locale.value;
	return (key: string, params?: Readonly<Record<string, string | number>>) =>
		params ? t(key, params as Record<string, unknown>) : t(key);
});

function change(event: Event): void {
	controller?.applyChartData((event as ChartFormatChangeEvent).detail.chartData);
}
function retarget(event: Event): void {
	controller?.setTarget((event as ChartFormatTargetEvent).detail.target);
}
</script>

<template>
	<pptx-ui-chart-format-pane
		v-if="visible && controller"
		:translateLabel.prop="translate"
		:chartData.prop="toRaw(controller.chart.value?.chartData)"
		:target.prop="controller.target.value"
		:disabled.prop="!controller.editable.value"
		@chart-format-change="change"
		@chart-format-target="retarget"
		@chart-format-close="controller.close()"
	/>
</template>
