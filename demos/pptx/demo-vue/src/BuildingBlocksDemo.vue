<script setup lang="ts">
import {
	BuildingBlocksCanvas,
	BuildingBlocksSlidesPane,
	BuildingBlocksToolbar,
	useViewerBuildingBlocks,
} from 'pptx-vue-viewer';
import { onMounted, ref, shallowRef, watchEffect } from 'vue';

import { mountBuildingBlocksPanel } from '../../shared/building-blocks-panel';

import 'pptx-vue-viewer/styles';

/**
 * `?blocks=1`: a custom shell composed from the Vue building blocks (toolbar, slides pane, canvas)
 * with the shared host panel that finishes their workflows. The framework-neutral
 * `e2e/pptx/building-blocks.spec.ts` drives it.
 */
const content = shallowRef<Uint8Array | null>(null);
const canEdit = ref(false);
const panel = ref<HTMLElement | null>(null);
const failure = ref('');
const blocks = useViewerBuildingBlocks(() => ({
	content: content.value,
	canEdit: canEdit.value,
	fileName: 'sample-deck.pptx',
}));

onMounted(async () => {
	try {
		const response = await fetch(`${import.meta.env.BASE_URL}sample-deck.pptx`);
		content.value = new Uint8Array(await response.arrayBuffer());
	} catch (error) {
		failure.value = String(error);
	}
	const host = mountBuildingBlocksPanel(panel.value!, {
		actions: () => blocks.actions.value,
		state: () => blocks.state.value,
		setCanEdit: (value) => {
			canEdit.value = value;
		},
	});
	watchEffect(() => {
		// Re-render the panel after every change of the blocks' actions or state.
		void blocks.actions.value;
		void blocks.state.value;
		host.refresh();
	});
});
</script>

<template>
	<p v-if="failure" role="alert">{{ failure }}</p>
	<section ref="panel" />
	<BuildingBlocksToolbar :blocks="blocks" />
	<div style="display: flex; height: calc(100vh - 160px)">
		<div style="width: 200px; overflow: auto">
			<BuildingBlocksSlidesPane :blocks="blocks" />
		</div>
		<div style="flex: 1; position: relative">
			<BuildingBlocksCanvas :blocks="blocks" />
		</div>
	</div>
</template>
