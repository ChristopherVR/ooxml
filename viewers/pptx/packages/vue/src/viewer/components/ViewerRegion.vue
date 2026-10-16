<script setup lang="ts">
/**
 * ViewerRegion: the one region host component. It renders a container `div`
 * and hands it to the shared `ViewerRegionHost`, which mounts the host's
 * renderer (a mount function or a custom element), delivers context changes
 * and tears it down. The host does the change detection, so this component
 * only forwards `renderer` and `context` whenever either changes.
 */
import { ViewerRegionHost } from 'ooxml-ui/pptx';
import type { ViewerRegionContext, ViewerRegionRenderer } from 'ooxml-ui/pptx';
import { onBeforeUnmount, onMounted, ref, watch } from 'vue';

const props = defineProps<{
	renderer: ViewerRegionRenderer;
	context: ViewerRegionContext;
}>();

const container = ref<HTMLElement | null>(null);
let host: ViewerRegionHost | undefined;

function render(): void {
	host?.render(props.renderer, props.context);
}

onMounted(() => {
	if (container.value) {
		host = new ViewerRegionHost(container.value);
		render();
	}
});
watch(() => [props.renderer, props.context] as const, render);
onBeforeUnmount(() => {
	host?.destroy();
	host = undefined;
});
</script>

<template>
	<div ref="container" class="pptx-vue-region" style="width: 100%; height: 100%"></div>
</template>
