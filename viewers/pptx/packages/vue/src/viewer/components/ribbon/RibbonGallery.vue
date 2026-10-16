<script setup lang="ts">
import type {
	RibbonControlId,
	RibbonGalleryId,
	RibbonGalleryPickEvent,
	RibbonGalleryPreviewEvent,
} from 'ooxml-ui/pptx';

import { useRibbonGallery } from './use-ribbon-gallery';

const props = defineProps<{
	gallery: RibbonGalleryId;
	control?: RibbonControlId;
	mode: 'inline' | 'dropdown' | 'chevron';
}>();
const emit = defineEmits<{ command: [command: 'formatBackground'] }>();
const g = useRibbonGallery(() => props.gallery);
function pick(event: Event): void {
	const { itemId, action } = (event as RibbonGalleryPickEvent).detail;
	const result = g.pick(itemId, action);
	if (result?.kind === 'command') {
		emit('command', result.command);
	}
}
function preview(event: Event): void {
	g.preview((event as RibbonGalleryPreviewEvent).detail.itemId);
}
</script>
<template>
	<pptx-ui-ribbon-gallery
		:data-ribbon-control="props.control"
		:mode="props.mode === 'inline' ? 'inline' : 'dropdown'"
		:chevron-only="props.mode === 'chevron' || undefined"
		:descriptor.prop="g.descriptor.value"
		:translateLabel.prop="g.translate"
		:disabled.prop="g.disabled.value"
		@gallery-pick="pick"
		@gallery-preview="preview"
		@gallery-preview-end="g.endPreview()"
	></pptx-ui-ribbon-gallery>
</template>
