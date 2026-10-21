<script setup lang="ts">
import type { OfficeColorPickEvent, ThemeColorPickerCommit } from 'ooxml-ui/pptx';
import { buildThemeColorSwatchGrid, themeColorGridCommit, themeColorGridView } from 'ooxml-ui/pptx';
import type { PptxThemeColorRef } from 'pptx-viewer-core';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

import { injectThemeColorMap } from '../../composables/theme-color-map-context';

/**
 * PowerPoint's "Theme Colors" grid on the shared `pptx-ui-color-grid`: ten columns (Background 1,
 * Text 1, Background 2, Text 2, Accent 1..6), each a base swatch over five luminance variants,
 * built from the loaded deck's theme (`ThemeColorMapKey`) rather than a hard-coded Office palette.
 *
 * Renders nothing when no deck theme is loaded yet, so callers can render this unconditionally
 * alongside their existing hex and recent-colour controls.
 */
const props = defineProps<{
	disabled?: boolean;
	/** The element's current theme ref, if any (highlights the matching swatch). */
	selectedRef?: PptxThemeColorRef;
	/** The element's current resolved hex, used to highlight a swatch when no ref is stored. */
	selectedHex?: string;
}>();

const emit = defineEmits<{
	pick: [commit: ThemeColorPickerCommit];
}>();

// The inspector keeps its own standard and recent rows, so the grid shows only Theme Colors.
const NO_STANDARD_COLORS: readonly never[] = [];

const { t } = useI18n();
const themeColorMap = injectThemeColorMap();

const columns = computed(() =>
	buildThemeColorSwatchGrid(themeColorMap?.value, (key, params) =>
		params ? t(key, { ...params }) : t(key),
	),
);
const view = computed(() =>
	themeColorGridView(columns.value, props.selectedRef, props.selectedHex),
);
const heading = computed(() => t('pptx.colorPicker.themeColors'));

function pick(event: Event): void {
	const commit = themeColorGridCommit(columns.value, (event as OfficeColorPickEvent).detail);
	if (commit) {
		emit('pick', commit);
	}
}
</script>

<template>
	<pptx-ui-color-grid
		v-if="columns.length > 0"
		class="pptx-vue-theme-swatch-grid mt-1"
		:palette.prop="view.palette"
		:selectedTheme.prop="view.selectedTheme"
		:standardColors.prop="NO_STANDARD_COLORS"
		:themeHeading.prop="heading"
		:label.prop="heading"
		:disabled.prop="disabled"
		@office-color-pick="pick"
	/>
</template>
