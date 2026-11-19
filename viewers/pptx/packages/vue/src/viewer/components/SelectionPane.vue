<script setup lang="ts">
/** The Selection Pane: rows, rename box and drag reorder from the shared selection pane controller. */
import { Eye, EyeOff, GripVertical } from '@lucide/vue';
import {
	createSelectionPaneController,
	formatSelectionPaneLabel,
	restoreEditorKeyboardFocus,
	selectionPaneRows,
} from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { computed, nextTick, ref } from 'vue';
import { useI18n } from 'vue-i18n';

import { useController } from '../composables/useController';

const props = defineProps<{
	elements: PptxElement[];
	selectedIds: string[];
	canEdit: boolean;
}>();

const emit = defineEmits<{
	select: [id: string];
	'toggle-visibility': [id: string];
	reorder: [payload: { from: number; to: number }];
	rename: [payload: { id: string; name: string }];
	close: [];
}>();

const { t } = useI18n();

const pane = createSelectionPaneController({
	canEdit: () => props.canEdit,
	onRename: (id, name) => emit('rename', { id, name }),
	onReorder: (from, to) => emit('reorder', { from, to }),
});
const state = useController(pane);
const editingId = computed(() => state.value.renamingId);
const dragOverIndex = computed(() => state.value.dragOverIndex);

const rows = computed(() =>
	selectionPaneRows(props.elements).map((row) => ({
		...row,
		name: formatSelectionPaneLabel(row.label, (key, params) => t(key, params ?? {})),
	})),
);

const renameInputRef = ref<HTMLInputElement | null>(null);
function setRenameInput(el: unknown): void {
	renameInputRef.value = el instanceof HTMLInputElement ? el : null;
}

function startRename(element: PptxElement, name: string): void {
	if (!pane.startRename(element.id, name)) {
		return;
	}
	void nextTick(() => {
		renameInputRef.value?.focus();
		renameInputRef.value?.select();
	});
}

/** Give the keyboard back to the viewer root before the input goes away. */
function releaseRenameFocus(): void {
	restoreEditorKeyboardFocus(renameInputRef.value);
}

// The box closes first, then the keyboard goes back to the root: focusing the root blurs the
// input, and the blur's own commit must find the box already closed.
function commitRename(): void {
	if (state.value.renamingId === null) {
		return;
	}
	pane.commitRename();
	releaseRenameFocus();
}

function cancelRename(): void {
	pane.cancelRename();
	releaseRenameFocus();
}

function onDragOver(event: DragEvent, index: number): void {
	event.preventDefault();
	pane.dragOver(index);
}
</script>

<template>
	<div class="flex h-full w-56 flex-col border-l border-border bg-popover" data-pptx-selection-pane>
		<div class="flex items-center justify-between border-b border-border px-3 py-2">
			<span class="text-xs font-medium text-foreground">{{ t('pptx.selectionPane.heading') }}</span>
			<button
				type="button"
				class="text-xs text-muted-foreground hover:text-foreground"
				:title="t('pptx.selectionPane.close')"
				@click="emit('close')"
			>
				&times;
			</button>
		</div>
		<div class="flex-1 overflow-y-auto py-1">
			<div v-if="rows.length === 0" class="px-3 py-4 text-xs italic text-muted-foreground">
				{{ t('pptx.selectionPane.noObjects') }}
			</div>
			<div
				v-for="{ element, index, name } in rows"
				v-else
				:key="element.id"
				:draggable="props.canEdit && editingId !== element.id"
				class="flex cursor-pointer items-center gap-1.5 px-2 py-1 text-xs transition-colors"
				:class="[
					props.selectedIds.includes(element.id)
						? 'bg-primary/30 text-primary'
						: 'text-foreground hover:bg-muted',
					dragOverIndex === index ? 'border-t-2 border-primary' : '',
				]"
				@click="emit('select', element.id)"
				@dragstart="pane.dragStart(index)"
				@dragover="onDragOver($event, index)"
				@drop="pane.drop(index)"
				@dragend="pane.dragEnd()"
			>
				<GripVertical
					v-if="props.canEdit"
					class="h-3 w-3 flex-shrink-0 cursor-grab text-muted-foreground"
				/>
				<span
					class="flex-1 truncate"
					data-pptx-selection-name
					@dblclick.stop="startRename(element, name)"
				>
					<input
						v-if="editingId === element.id"
						:ref="setRenameInput"
						type="text"
						:aria-label="t('pptx.selectionPane.renameElement')"
						:value="state.renameDraft"
						class="w-full rounded border border-border bg-muted px-1 py-0.5 text-xs outline-none"
						@input="pane.setRenameDraft(($event.target as HTMLInputElement).value)"
						@blur="commitRename()"
						@keydown.enter="commitRename()"
						@keydown.escape="cancelRename"
						@click.stop
					/>
					<template v-else>{{ name }}</template>
				</span>
				<button
					type="button"
					class="flex-shrink-0 text-muted-foreground hover:text-foreground"
					:title="element.hidden ? t('pptx.selectionPane.show') : t('pptx.selectionPane.hide')"
					@click.stop="emit('toggle-visibility', element.id)"
				>
					<EyeOff v-if="element.hidden" class="h-3.5 w-3.5" />
					<Eye v-else class="h-3.5 w-3.5 opacity-50" />
				</button>
			</div>
		</div>
	</div>
</template>
