<script setup lang="ts">
/**
 * RecentColorsRow: the "Recent colours" row shown under a colour picker once
 * at least one colour has been used (`p:clrMru`, seeded and kept by
 * `useRecentColors`), drawn by the shared `pptx-ui-color-grid` with only its
 * Recent Colors section (`recentColorsGridView`). Presentational only:
 * clicking a swatch just emits `pick`, the host applies it exactly like a
 * picked-from-the-native-picker colour (which folds it back to the front of
 * the list too).
 */
import type { OfficeColorPickEvent } from 'ooxml-ui/pptx';
import { recentColorsGridView } from 'ooxml-ui/pptx';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

const props = defineProps<{
	colors: string[];
	disabled?: boolean;
}>();

const emit = defineEmits<{
	pick: [hex: string];
}>();

const { t } = useI18n();
const view = computed(() =>
	recentColorsGridView(props.colors, (key, params) => (params ? t(key, { ...params }) : t(key))),
);

function pick(event: Event): void {
	emit('pick', (event as OfficeColorPickEvent).detail.color);
}
</script>

<template>
	<pptx-ui-color-grid
		v-if="colors.length > 0"
		class="pptx-vue-recent-colors mt-1"
		data-testid="pptx-color-recent"
		:palette.prop="view.palette"
		:standardColors.prop="view.standardColors"
		:recentColors.prop="view.recentColors"
		:recentHeading.prop="view.recentHeading"
		:recentLabel.prop="view.recentLabel"
		:label.prop="view.label"
		:disabled.prop="disabled"
		@office-color-pick="pick"
	/>
</template>
