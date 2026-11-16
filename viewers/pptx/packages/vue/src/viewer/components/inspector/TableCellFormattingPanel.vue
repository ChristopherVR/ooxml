<script setup lang="ts">
import {
	computeMergeCellDown,
	computeMergeCellRight,
	computeSplitCell,
	TABLE_CELL_SECTIONS,
	withCellStyle,
} from 'ooxml-ui/pptx';
import type { PptxTableCellStyle, PptxTableData } from 'pptx-viewer-core';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

import InspectorFields from './InspectorFields.vue';
import TableCellAdvancedFill from './TableCellAdvancedFill.vue';
import TableCellColorField from './TableCellColorField.vue';

/**
 * TableCellFormattingPanel: Vue port of React's inspector
 * `TableCellFormattingPanel.tsx`. Formats the currently-selected table cell:
 * the shared `TABLE_CELL_SECTIONS` (size, emphasis, alignment, margins, edge
 * and diagonal borders) drawn by the generic renderer, the text / background
 * colour pickers, advanced fill (gradient / pattern), and merge / split. All edits
 * emit a full `Partial<PptxTableData>` (rows) patch that the parent forwards to
 * the editor operations, so undo / redo and dirty-marking flow uniformly.
 */
const props = defineProps<{
	tableData: PptxTableData;
	rowIndex: number;
	columnIndex: number;
	canEdit: boolean;
}>();

const emit = defineEmits<{
	update: [patch: Partial<PptxTableData>];
}>();

const { t } = useI18n();

const cell = computed(() => props.tableData.rows[props.rowIndex]?.cells[props.columnIndex]);
const cs = computed<PptxTableCellStyle>(() => cell.value?.style ?? {});

function updateCellStyle(updates: Partial<PptxTableCellStyle>): void {
	const newRows = props.tableData.rows.map((row, ri) => {
		if (ri !== props.rowIndex) {
			return row;
		}
		return {
			...row,
			cells: row.cells.map((c, ci) => (ci !== props.columnIndex ? c : withCellStyle(c, updates))),
		};
	});
	emit('update', { rows: newRows });
}

function mergeRight(): void {
	const rows = computeMergeCellRight(props.tableData, props.rowIndex, props.columnIndex);
	if (rows) {
		emit('update', { rows });
	}
}

function mergeDown(): void {
	const rows = computeMergeCellDown(props.tableData, props.rowIndex, props.columnIndex);
	if (rows) {
		emit('update', { rows });
	}
}

function split(): void {
	const rows = computeSplitCell(props.tableData, props.rowIndex, props.columnIndex);
	if (rows) {
		emit('update', { rows });
	}
}
</script>

<template>
	<div v-if="cell" class="flex flex-col gap-2">
		<div class="text-[11px] font-semibold uppercase tracking-wide text-muted-foreground">
			{{ t('pptx.table.cell', { row: rowIndex + 1, col: columnIndex + 1 }) }}
		</div>

		<div class="grid grid-cols-2 gap-1.5">
			<TableCellColorField
				:label="t('pptx.table.color')"
				:value="cs.color"
				fallback="#000000"
				:selected-ref="cs.colorRef"
				:disabled="!canEdit"
				:aria-label="t('pptx.tableCell.textColorAria')"
				@commit="(hex, ref) => updateCellStyle({ color: hex, colorRef: ref })"
			/>
			<TableCellColorField
				:label="t('pptx.table.background')"
				:value="cs.backgroundColor"
				fallback="#ffffff"
				:selected-ref="cs.backgroundColorRef"
				:disabled="!canEdit"
				:aria-label="t('pptx.tableCell.backgroundColorAria')"
				@commit="(hex, ref) => updateCellStyle({ backgroundColor: hex, backgroundColorRef: ref })"
			/>
		</div>

		<TableCellAdvancedFill :cell-style="cs" :can-edit="canEdit" @update="updateCellStyle" />

		<InspectorFields
			:sections="TABLE_CELL_SECTIONS"
			:target="cs"
			:disabled="!canEdit"
			@patch="updateCellStyle"
		/>

		<div class="grid grid-cols-3 gap-1">
			<button
				type="button"
				class="rounded border border-border bg-muted px-2 py-1 text-center text-[11px] hover:bg-accent disabled:opacity-50 max-md:min-h-[44px]"
				:disabled="!canEdit"
				@click="mergeRight"
			>
				{{ t('pptx.table.mergeRight') }}
			</button>
			<button
				type="button"
				class="rounded border border-border bg-muted px-2 py-1 text-center text-[11px] hover:bg-accent disabled:opacity-50 max-md:min-h-[44px]"
				:disabled="!canEdit"
				@click="mergeDown"
			>
				{{ t('pptx.table.mergeDown') }}
			</button>
			<button
				type="button"
				class="rounded border border-border bg-muted px-2 py-1 text-center text-[11px] hover:bg-accent disabled:opacity-50 max-md:min-h-[44px]"
				:disabled="!canEdit"
				@click="split"
			>
				{{ t('pptx.table.split') }}
			</button>
		</div>
	</div>
</template>
