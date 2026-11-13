<script setup lang="ts">
/**
 * BuildingBlocksToolbar: the compact toolbar block, drawn from `describeBuildingBlockToolbar`.
 * Insert controls raise host requests; the rest run the viewer's own history, zoom and slides.
 */
import { describeBuildingBlockToolbar, runBuildingBlockToolbarControl } from 'ooxml-ui/pptx';
import type { BuildingBlockToolbarControlId, BuildingBlockToolbarGroupId } from 'ooxml-ui/pptx';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

import type { ViewerBuildingBlocks } from './types';

const props = defineProps<{
	blocks: ViewerBuildingBlocks;
	/** Render only these groups. */
	groups?: readonly BuildingBlockToolbarGroupId[];
}>();
const { t } = useI18n();

const described = computed(() => {
	const { state, actions, options } = props.blocks;
	const groups = describeBuildingBlockToolbar({
		availability: actions.value.availability,
		canUndo: state.value.canUndo,
		canRedo: state.value.canRedo,
		findOpen: actions.value.findReplace.open,
		slideIndex: state.value.slideIndex,
		slideCount: state.value.slideCount,
		hiddenActions: options.value.hiddenActions,
	});
	const include = props.groups;
	return include ? groups.filter((group) => include.includes(group.id)) : groups;
});

function run(id: BuildingBlockToolbarControlId): void {
	const view = props.blocks.bridge.engine.value?.view;
	runBuildingBlockToolbarControl(id, props.blocks.session, {
		undo: () => view?.undo(),
		redo: () => view?.redo(),
		toggleFind: () => view?.toggleFind(),
		zoomOut: () => view?.zoomOut(),
		zoomIn: () => view?.zoomIn(),
		zoomToFit: () => view?.zoomToFit(),
		previousSlide: () => view?.previousSlide(),
		nextSlide: () => view?.nextSlide(),
	});
}
</script>

<template>
	<div data-pptx-block="toolbar" class="pptx-vue-bb-toolbar" role="toolbar">
		<span v-for="group in described" :key="group.id" :data-group="group.id" role="group">
			<button
				v-for="control in group.controls"
				:key="control.id"
				type="button"
				:data-control="control.id"
				:disabled="control.disabled"
				:aria-pressed="control.pressed"
				@click="run(control.id)"
			>
				{{ t(control.labelKey) }}
			</button>
		</span>
	</div>
</template>

<style>
.pptx-vue-bb-toolbar {
	display: flex;
	flex-wrap: wrap;
	gap: 8px;
	padding: 4px;
}
.pptx-vue-bb-toolbar [data-group] {
	display: inline-flex;
	gap: 4px;
}
</style>
