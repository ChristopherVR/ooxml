<script setup lang="ts">
/**
 * MediaPlaybackOptions: the "Playback" card of the media inspector, the shared
 * `MEDIA_PLAYBACK_SECTIONS` (volume, start, play across slides, loop, full
 * screen, hide while not playing, fades and speed) drawn by the generic
 * renderer. Emits `update` with the shallow media patch the descriptor wrote.
 */
import { MEDIA_PLAYBACK_SECTIONS } from 'ooxml-ui/pptx';
import type { MediaPptxElement, PptxElement } from 'pptx-viewer-core';
import { useI18n } from 'vue-i18n';

import InspectorFields from './InspectorFields.vue';

const props = defineProps<{
	element: MediaPptxElement;
	canEdit: boolean;
}>();

const emit = defineEmits<{
	update: [patch: Partial<PptxElement>];
}>();

const { t } = useI18n();
</script>

<template>
	<div class="pptx-vue-media-card rounded border border-border bg-card p-2 space-y-2">
		<div class="text-[11px] uppercase tracking-wide text-muted-foreground">
			{{ t('pptx.media.playback') }}
		</div>
		<InspectorFields
			:sections="MEDIA_PLAYBACK_SECTIONS"
			:target="props.element as PptxElement"
			:disabled="!props.canEdit"
			@patch="emit('update', $event)"
		/>
	</div>
</template>
