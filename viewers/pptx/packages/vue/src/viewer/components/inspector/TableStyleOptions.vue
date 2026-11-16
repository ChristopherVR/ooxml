<script setup lang="ts">
import {
	applyTableStylePreset,
	TABLE_OPTIONS_SECTIONS,
	TABLE_STYLE_PRESETS,
	tableStyleAssignmentUpdate,
} from 'ooxml-ui/pptx';
import type { ParsedTableStyleMap, PptxTableData } from 'pptx-viewer-core';
import { ref } from 'vue';
import { useI18n } from 'vue-i18n';

import InspectorFields from './InspectorFields.vue';
import TableStyleEditor from './TableStyleEditor.vue';

/**
 * TableStyleOptions: the shared table options (`TABLE_OPTIONS_SECTIONS`, drawn
 * by the generic renderer) and the quick-style preset grid. Emits a
 * `Partial<PptxTableData>` patch (banding flags, band cycle counts, or a full
 * re-styled `rows` array) that the parent merges into the element's tableData.
 *
 * `tableStyleMap` is optional/absent when the host has not yet wired the
 * table-style-DEFINITION-editor feature through (see `TableStyleEditor.vue`'s
 * docblock); the "Edit style..." button then simply does not render.
 */
const props = defineProps<{
	tableData: PptxTableData;
	canEdit: boolean;
	tableStyleMap?: ParsedTableStyleMap;
}>();

const emit = defineEmits<{
	update: [patch: Partial<PptxTableData>];
	tableStyleMapChange: [nextMap: ParsedTableStyleMap];
	deleteTableStyle: [styleId: string];
}>();

const showStyleEditor = ref(false);

const { t } = useI18n();

function applyPreset(preset: (typeof TABLE_STYLE_PRESETS)[number]): void {
	emit('update', { rows: applyTableStylePreset(props.tableData, preset) });
}

/** A newly-created style (from "Edit style...") becomes this table's style. */
function assignStyle(styleId: string): void {
	emit('update', tableStyleAssignmentUpdate(styleId));
}
</script>

<template>
	<div class="flex flex-col gap-2">
		<InspectorFields
			:sections="TABLE_OPTIONS_SECTIONS"
			:target="tableData"
			:disabled="!canEdit"
			@patch="emit('update', $event)"
		/>

		<div class="flex flex-col gap-1">
			<span class="text-[11px] font-medium">{{ t('pptx.table.stylePresets') }}</span>
			<div class="grid grid-cols-3 gap-1.5">
				<button
					v-for="preset in TABLE_STYLE_PRESETS"
					:key="preset.id"
					type="button"
					class="h-10 overflow-hidden rounded border border-border transition-colors hover:border-primary disabled:opacity-50"
					:disabled="!canEdit"
					:title="t(preset.labelKey)"
					@click="applyPreset(preset)"
				>
					<div class="flex h-full flex-col">
						<div class="flex-1" :style="{ backgroundColor: preset.headerBg }" />
						<div class="flex-1" :style="{ backgroundColor: preset.bandBg }" />
						<div class="flex-1 border-t" :style="{ borderColor: preset.borderColor }" />
					</div>
				</button>
			</div>
			<button
				v-if="tableStyleMap !== undefined"
				type="button"
				class="mt-1.5 self-start rounded bg-muted hover:bg-accent px-2 py-1 text-[11px] transition-colors"
				:disabled="!canEdit"
				@click="showStyleEditor = !showStyleEditor"
			>
				{{ t('pptx.tableStyleEditor.editButton') }}
			</button>
		</div>

		<TableStyleEditor
			v-if="showStyleEditor && tableStyleMap !== undefined"
			:style-map="tableStyleMap"
			:style-id="tableData.tableStyleId"
			:can-edit="canEdit"
			@style-map-change="(m) => emit('tableStyleMapChange', m)"
			@delete-style="(id) => emit('deleteTableStyle', id)"
			@assign-style="assignStyle"
			@close="showStyleEditor = false"
		/>
	</div>
</template>
