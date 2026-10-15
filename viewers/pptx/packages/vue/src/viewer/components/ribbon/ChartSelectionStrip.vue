<script setup lang="ts">
/**
 * ChartSelectionStrip: Chart Format > Current Selection, the shared
 * `pptx-ui-ribbon-home-chart-selection` strip (element drop-down, Format
 * Selection, Reset to Match Style) over the injected chart-format state.
 */
import type { RibbonHomeRequestEvent } from 'ooxml-ui/pptx';
import {
	chartSelectionHomeControls,
	dispatchChartSelectionRequest,
	homeSnapshotTranslator,
} from 'ooxml-ui/pptx';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

import { injectChartFormat } from '../../composables/chart-format-pane';

const { t, locale } = useI18n();
const controller = injectChartFormat();

const state = computed(() => ({
	controls: chartSelectionHomeControls({
		chartData: controller?.chart.value?.chartData,
		targetKey: controller?.target.value,
		editable: Boolean(controller?.editable.value),
		translate: (key, params) => (params ? t(key, params) : t(key)),
	}),
	// The locale is read so a language switch re-translates the shared labels.
	locale: locale.value,
	translate: homeSnapshotTranslator(['chart-selection'], t),
}));

function request(event: RibbonHomeRequestEvent): void {
	if (!controller) {
		return;
	}
	dispatchChartSelectionRequest(event.detail, {
		select: controller.setTarget,
		formatSelection: () => controller.openPane(),
		reset: controller.reset,
	});
}
</script>

<template>
	<pptx-ui-ribbon-home-chart-selection :state.prop="state" @home-request="request" />
</template>
