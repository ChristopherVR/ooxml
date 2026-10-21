<script setup lang="ts">
import type { ExportProgressViewState } from 'ooxml-ui/pptx';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

/**
 * ExportProgressModal - the progress modal shown while a multi-slide export (PDF / GIF / WebM)
 * runs. A thin adapter over the shared `pptx-ui-export-progress`, which owns the dialog and
 * progress bar semantics, focus and the look, and is deliberately not dismissable: the export
 * ends by completing, erroring or Cancel. Cancel emits `cancel`; the parent aborts the export
 * and clears `open`.
 *
 * Rendered in place (not teleported to `body`), so it takes the viewer's theme variables, and
 * kept mounted while closed, so it hands focus back when an export ends.
 */
const props = defineProps<{
	/** Whether the overlay is visible. */
	open: boolean;
	/** Heading shown at the top (e.g. "Export as PDF"). */
	title: string;
	/** Current progress, 0-100. */
	progress: number;
	/** Optional status line (e.g. "Rendering slide 3 of 10..."). */
	statusMessage?: string;
}>();

const emit = defineEmits<{
	(e: 'cancel'): void;
}>();

const { t, locale } = useI18n();

const state = computed<ExportProgressViewState>(() => {
	void locale.value;
	return {
		open: props.open,
		title: props.title,
		progress: props.progress,
		status: props.statusMessage,
		translate: (key) => t(key),
	};
});
</script>

<template>
	<pptx-ui-export-progress :state.prop="state" @export-cancel="emit('cancel')" />
</template>
