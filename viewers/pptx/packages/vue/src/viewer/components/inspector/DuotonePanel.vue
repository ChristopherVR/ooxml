<script setup lang="ts">
import { DUOTONE_PRESETS, pictureDuotonePresetPatch } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { useI18n } from 'vue-i18n';

/**
 * DuotonePanel: the duotone preset swatches, the native half of the picture's
 * Duotone effect. The toggle and its two colours are the shared
 * `PICTURE_DUOTONE_SECTION`; a swatch applies the shared
 * `pictureDuotonePresetPatch`.
 */
const props = defineProps<{
	element: PptxElement;
	canEdit?: boolean;
}>();

const emit = defineEmits<{
	update: [patch: Partial<PptxElement>];
}>();

const { t } = useI18n();
</script>

<template>
	<div class="pptx-vue-duotone flex flex-col gap-1 text-[11px]">
		<span class="text-[10px] text-muted-foreground">{{ t('pptx.image.duotonePresets') }}</span>
		<div class="pptx-vue-duotone__grid grid grid-cols-4 gap-1">
			<button
				v-for="preset in DUOTONE_PRESETS"
				:key="preset.labelKey"
				type="button"
				data-testid="picture-duotone-preset"
				class="pptx-vue-duotone__cell flex flex-col items-center gap-0.5 rounded border border-border hover:bg-accent p-1 transition-colors"
				:disabled="props.canEdit === false"
				:title="t(preset.labelKey)"
				@click="emit('update', pictureDuotonePresetPatch(props.element, preset))"
			>
				<div class="flex h-4 w-full overflow-hidden rounded">
					<div class="flex-1" :style="{ backgroundColor: preset.shadow }" />
					<div class="flex-1" :style="{ backgroundColor: preset.highlight }" />
				</div>
				<span class="w-full truncate text-center text-[8px] text-muted-foreground">
					{{ t(preset.labelKey) }}
				</span>
			</button>
		</div>
	</div>
</template>
