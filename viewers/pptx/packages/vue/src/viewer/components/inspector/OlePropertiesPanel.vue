<script setup lang="ts">
import { OLE_OBJECT_NAME_SECTIONS, resolveOleSummaryRows } from 'ooxml-ui/pptx';
/**
 * OlePropertiesPanel: OLE object summary (type / file name / link status)
 * plus the Object Name editor, at parity with React's `ElementMiscPanels.tsx`
 * OlePropertiesPanel.
 *
 * A browser cannot run the native application that owns an embedded OLE
 * object, so the object itself stays read-only. Its Object Name IS editable:
 * `p:oleObj/@name` (ECMA-376 SS13.3.4) already parses, saves, and syncs via
 * collaboration, and shared's `getOleDisplayName` / `getOleAriaLabel` already
 * read it, so this field was the only piece missing to make it a real,
 * round-tripping edit.
 */
import type { OlePptxElement, PptxElement } from 'pptx-viewer-core';
import { computed, ref } from 'vue';
import { useI18n } from 'vue-i18n';

import OleEditorDialog from '../OleEditorDialog.vue';
import InspectorFields from './InspectorFields.vue';

const props = defineProps<{ element: PptxElement; canEdit?: boolean }>();
const emit = defineEmits<{
	update: [patch: Partial<PptxElement>];
}>();
const { t } = useI18n();

const ole = computed(() => props.element as OlePptxElement);
const isEditorOpen = ref(false);

const summaryRows = computed(() => resolveOleSummaryRows(props.element));

function onEditorUpdate(patch: Partial<PptxElement>): void {
	emit('update', patch);
}
</script>

<template>
	<div class="space-y-1.5 text-[11px]">
		<button
			v-if="props.canEdit && !ole.isLinked"
			type="button"
			class="w-full rounded border border-border px-2 py-1.5 text-[11px] text-foreground transition-colors hover:bg-accent"
			@click="isEditorOpen = true"
		>
			{{ t('pptx.ole.editContent') }}
		</button>
		<InspectorFields
			:sections="OLE_OBJECT_NAME_SECTIONS"
			:target="props.element"
			:disabled="!props.canEdit"
			@patch="emit('update', $event)"
		/>
		<div v-for="row in summaryRows" :key="row.id" class="flex items-center justify-between gap-2">
			<span class="text-muted-foreground">{{ t(row.labelKey) }}</span>
			<span
				v-if="row.id === 'linkStatus'"
				class="inline-flex items-center rounded-full px-2 py-0.5 text-[10px] font-medium"
				:class="ole.isLinked ? 'bg-blue-500/20 text-blue-400' : 'bg-green-500/20 text-green-400'"
			>
				{{ t(row.valueKey ?? '') }}
			</span>
			<span v-else class="truncate text-foreground" :title="row.value">{{ row.value }}</span>
		</div>
	</div>
	<OleEditorDialog
		:open="isEditorOpen"
		:element="ole"
		@update="onEditorUpdate"
		@close="isEditorOpen = false"
	/>
</template>
