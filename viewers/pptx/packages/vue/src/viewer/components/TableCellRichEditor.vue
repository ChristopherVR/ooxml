<script setup lang="ts">
/**
 * Rich inline editor for a table cell being edited: mounts the shared
 * `openTableCellTextEditor` (the editor every binding uses), so the cell's runs
 * keep their formatting, Enter adds a paragraph and Tab, the arrows at the
 * text's edge and Escape leave the cell. The end of the edit is emitted once;
 * unmounting an editor that is still open reports a `detach` end, which commits.
 */
import type {
	TableCellCaret,
	TableCellCss,
	TableCellEditEnd,
	TableCellTextEditorSession,
} from 'ooxml-ui/pptx';
import { openTableCellTextEditor, tableCellEditorSegments } from 'ooxml-ui/pptx';
import type { PptxTableCell, TablePptxElement } from 'pptx-viewer-core';
import { onBeforeUnmount, onMounted, ref } from 'vue';

const props = defineProps<{
	cell: PptxTableCell | undefined;
	/** The table and the cell's position, so the editor keeps each run's markup. */
	element?: TablePptxElement;
	row?: number;
	col?: number;
	cellCss?: TableCellCss;
	caret?: TableCellCaret;
	rtl?: boolean;
}>();
const emit = defineEmits<{ end: [end: TableCellEditEnd] }>();

const host = ref<HTMLDivElement | null>(null);
let session: TableCellTextEditorSession | undefined;

onMounted(() => {
	if (!host.value) {
		return;
	}
	session = openTableCellTextEditor({
		host: host.value,
		segments: tableCellEditorSegments(props.cell, props.element, props.row, props.col),
		...(props.cellCss ? { cellCss: props.cellCss } : {}),
		...(props.caret ? { caret: props.caret } : {}),
		rtl: props.rtl === true,
		onEnd: (end) => emit('end', end),
	});
});

onBeforeUnmount(() => {
	session?.destroy();
	session = undefined;
});
</script>

<template>
	<div
		ref="host"
		class="pptx-vue-table__cell-editor"
		@pointerdown.stop
		@mousedown.stop
		@click.stop
		@dblclick.stop
	/>
</template>
