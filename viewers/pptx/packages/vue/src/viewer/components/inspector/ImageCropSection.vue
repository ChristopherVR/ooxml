<script setup lang="ts">
import { cropResetPatch, cropResetState, PICTURE_CROP_SECTIONS } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

import InspectorFields from './InspectorFields.vue';

/**
 * ImageCropSection: the four crop insets (the shared `PICTURE_CROP_SECTIONS`
 * drawn by the generic renderer), a Reset Crop action, and a "Replace Image"
 * button. Crop insets are top-level element fields (fractions 0..1).
 *
 * Replace Image reads the chosen file as a data-URL and emits `{ imageData }`
 * so the swap happens entirely within the panel (no host wiring required).
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

// G7: `a:picLocks/@noCrop` forbids cropping this specific picture.
const croppable = computed(() => cropResetState(props.element, props.canEdit, true).enabled);

function onCrop(patch: Partial<PptxElement>): void {
	if (croppable.value) {
		emit('update', patch);
	}
}

function onResetCrop(): void {
	if (!croppable.value) {
		return;
	}
	emit('update', cropResetPatch());
}

function onReplaceImage(event: Event): void {
	const input = event.target as HTMLInputElement;
	const file = input.files?.[0];
	input.value = '';
	if (!file) {
		return;
	}
	const reader = new FileReader();
	reader.onload = () => {
		if (typeof reader.result === 'string') {
			emit('update', { imageData: reader.result } as Partial<PptxElement>);
		}
	};
	reader.readAsDataURL(file);
}
</script>

<template>
	<div class="pptx-vue-image-crop flex flex-col gap-2 text-[11px]">
		<label
			class="pptx-vue-image-crop__replace w-full text-center rounded border border-border bg-muted hover:bg-accent px-2 py-1 cursor-pointer transition-colors"
		>
			{{ t('pptx.image.replaceImage') }}
			<input
				type="file"
				class="hidden"
				accept="image/png,image/jpeg,image/gif,image/webp,image/svg+xml"
				@change="onReplaceImage"
			/>
		</label>

		<InspectorFields
			:sections="PICTURE_CROP_SECTIONS"
			:target="props.element"
			:disabled="!croppable"
			@patch="onCrop"
		/>

		<button
			type="button"
			class="pptx-vue-image-crop__reset w-full rounded border border-border bg-muted hover:bg-accent px-2 py-1 transition-colors"
			:disabled="!croppable"
			@click="onResetCrop"
		>
			{{ t('pptx.image.resetCrop') }}
		</button>
	</div>
</template>
