<script setup lang="ts">
/**
 * InsertChartDialog: Insert > Chart > Insert Chart..., the shared
 * `pptx-ui-insert-chart-dialog` (type gallery with previews, grouping, title). The element owns
 * the dialog; this component maps its open state and routes the choice to `onAddChart`.
 */
import type {
	InsertChartDialogInsertEvent,
	InsertChartKind,
	InsertChartOptions,
} from 'ooxml-ui/pptx';
import { useI18n } from 'vue-i18n';

const props = defineProps<{
	open: boolean;
	onAddChart?: (chartKind: InsertChartKind, options?: InsertChartOptions) => void;
}>();
const emit = defineEmits<{ close: [] }>();
const { t } = useI18n();
const translateLabel = (key: string): string => t(key);

function insert(event: Event): void {
	const { kind, title, grouping } = (event as InsertChartDialogInsertEvent).detail;
	props.onAddChart?.(kind, { title, grouping });
}
</script>

<template>
	<pptx-ui-insert-chart-dialog
		:open.prop="props.open"
		:translate-label.prop="translateLabel"
		@insert-chart="insert"
		@dialog-close="emit('close')"
	/>
</template>
