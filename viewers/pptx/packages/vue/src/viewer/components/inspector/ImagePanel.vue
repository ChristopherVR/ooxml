<script setup lang="ts">
import {
	ELEMENT_ALT_TEXT_SECTIONS,
	imageResetPatch,
	imageResetState,
	PICTURE_ADJUSTMENT_SECTIONS,
	PICTURE_ARTISTIC_SECTIONS,
	PICTURE_EFFECT_SECTIONS,
} from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { isImageLikeElement } from 'pptx-viewer-core';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

import DuotonePanel from './DuotonePanel.vue';
import ImageCropSection from './ImageCropSection.vue';
import InspectorFields from './InspectorFields.vue';

/**
 * ImagePanel: inspector panel for image-like elements (`image` / `picture`),
 * at parity with React's ImagePropertiesPanel. The alt text and the picture
 * adjustments are the shared descriptors (`ELEMENT_ALT_TEXT_SECTIONS`,
 * `PICTURE_ADJUSTMENT_SECTIONS`) drawn by the generic renderer; the compound
 * controls are focused sub-panels (crop/replace, recolour,
 * colour wash, duotone). Each sub-panel emits a shallow `Partial<PptxElement>`
 * patch (nested `imageEffects` emitted whole) which this panel relays upward;
 * the host merges it via `ops.updateElement(id, patch)`.
 */
const props = withDefaults(
	defineProps<{
		element: PptxElement;
		canEdit?: boolean;
	}>(),
	{ canEdit: true },
);

const emit = defineEmits<{
	update: [patch: Partial<PptxElement>];
}>();

const { t } = useI18n();

const isImage = computed(() => isImageLikeElement(props.element));

/** Shared gating: enabled only while editable and an effect or crop shape is set. */
const resetState = computed(() => imageResetState(props.element, props.canEdit, true));

function relay(patch: Partial<PptxElement>): void {
	emit('update', patch);
}

/** Clear every adjustment, effect, and crop back to the picture's default. */
function onResetPicture(): void {
	emit('update', imageResetPatch());
}
</script>

<template>
	<div class="pptx-vue-image-panel flex flex-col gap-3 text-[11px]">
		<p v-if="!isImage" class="pptx-vue-image-panel__note text-muted-foreground">
			{{ t('pptx.image.noImageProperties') }}
		</p>

		<template v-else>
			<InspectorFields
				:sections="ELEMENT_ALT_TEXT_SECTIONS"
				:target="element"
				:disabled="!canEdit"
				@patch="relay"
			/>

			<ImageCropSection :element="element" :can-edit="canEdit" @update="relay" />

			<InspectorFields
				:sections="PICTURE_ADJUSTMENT_SECTIONS"
				:target="element"
				:disabled="!canEdit"
				@patch="relay"
			/>

			<InspectorFields
				:sections="PICTURE_ARTISTIC_SECTIONS"
				:target="element"
				:disabled="!canEdit"
				@patch="relay"
			/>

			<InspectorFields
				:sections="PICTURE_EFFECT_SECTIONS"
				:target="element"
				:disabled="!canEdit"
				@patch="relay"
			/>

			<DuotonePanel :element="element" :can-edit="canEdit" @update="relay" />

			<button
				type="button"
				:disabled="!resetState.enabled"
				class="pptx-vue-image-panel__reset-picture w-full rounded border border-border bg-muted hover:bg-accent px-2 py-1 transition-colors"
				@click="onResetPicture"
			>
				{{ t('pptx.image.resetImage') }}
			</button>
		</template>
	</div>
</template>
