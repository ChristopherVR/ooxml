<script setup lang="ts">
import { buildQuickStylesGallery, quickStyleShapeStyle } from 'ooxml-ui/pptx';
import type { RibbonGalleryPickEvent } from 'ooxml-ui/pptx';
import type { ShapeStyle } from 'pptx-viewer-core';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

/**
 * QuickStylesGallery: the inspector's Quick Styles, the shared
 * `pptx-ui-ribbon-gallery` panel (compact) over `buildQuickStylesGallery`, so
 * the presets, previews, accessible names and touch sizes match every other
 * binding.
 *
 * A pick emits `select` with the preset's `Partial<ShapeStyle>`
 * (`quickStyleShapeStyle`). The parent (`EffectsPanel`) owns the current
 * `shapeStyle` and merges the preset onto it, then forwards the FULL merged
 * sub-object via its own `update` patch, so this component stays pure
 * presentation.
 */
const props = defineProps<{
	/** Read-only: the tiles stay visible but cannot be picked. */
	disabled?: boolean;
}>();

const emit = defineEmits<{
	select: [style: Partial<ShapeStyle>];
}>();

const { t, locale } = useI18n();

const descriptor = computed(() => buildQuickStylesGallery(props.disabled));

// A new translator identity per locale makes the element re-read its labels.
const translate = computed(() => {
	void locale.value;
	return (key: string, params?: Readonly<Record<string, string | number>>) =>
		params ? t(key, params as Record<string, unknown>) : t(key);
});

function pick(event: Event): void {
	const style = quickStyleShapeStyle((event as RibbonGalleryPickEvent).detail.itemId);
	if (style) {
		emit('select', style);
	}
}
</script>

<template>
	<div class="pptx-vue-quickstyles flex flex-col gap-1">
		<span class="pptx-vue-quickstyles-label text-muted-foreground">{{
			t('pptx.shape.quickStyles')
		}}</span>
		<pptx-ui-ribbon-gallery
			mode="panel"
			compact
			:descriptor.prop="descriptor"
			:translateLabel.prop="translate"
			:disabled.prop="props.disabled"
			@gallery-pick="pick"
		></pptx-ui-ribbon-gallery>
	</div>
</template>
