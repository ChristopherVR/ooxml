<script setup lang="ts">
import { buildTextWarpGallery, TEXT_WARP_PRESETS, textWarpStyle } from 'ooxml-ui/pptx';
import type { RibbonGalleryPickEvent } from 'ooxml-ui/pptx';
import type { TextStyle } from 'pptx-viewer-core';
import { computed, ref } from 'vue';
import { useI18n } from 'vue-i18n';

/**
 * TextWarpGallery: collapsible Text Warp (WordArt Transform) picker. The tiles
 * are the shared `pptx-ui-ribbon-gallery` panel over `buildTextWarpGallery`,
 * so the catalogue, previews, current-preset mark and accessible names match
 * every other binding. A pick emits `select` with `textWarpStyle(itemId)` (the
 * preset, or none for No Transform); the parent (`TextPanel`) merges it into
 * the full `textStyle` sub-object.
 *
 * The gallery stays mounted while collapsed (only hidden), so its tiles are
 * always there to query and expanding never rebuilds them.
 */
const props = defineProps<{
	ts: TextStyle | undefined;
	/** Read-only: the tiles show the current warp but cannot be picked. */
	disabled?: boolean;
}>();

const emit = defineEmits<{
	select: [style: Pick<TextStyle, 'textWarpPreset'>];
}>();

const { t, locale } = useI18n();

const expanded = ref(false);

const descriptor = computed(() => buildTextWarpGallery(props.ts?.textWarpPreset, props.disabled));

// A new translator identity per locale makes the element re-read its labels.
const translate = computed(() => {
	void locale.value;
	return (key: string, params?: Readonly<Record<string, string | number>>) =>
		params ? t(key, params as Record<string, unknown>) : t(key);
});

const currentLabel = computed<string>(() => {
	const current = props.ts?.textWarpPreset ?? 'textNoShape';
	const match = TEXT_WARP_PRESETS.find((p) => p.value === current);
	return match ? t(match.labelKey) : current;
});

function pick(event: Event): void {
	const style = textWarpStyle((event as RibbonGalleryPickEvent).detail.itemId);
	if (style) {
		emit('select', style);
	}
}
</script>

<template>
	<div class="pptx-vue-textwarp mt-2 rounded border border-border bg-card p-2 space-y-2">
		<button
			type="button"
			class="pptx-vue-textwarp-toggle flex w-full items-center justify-between text-[11px] uppercase tracking-wide text-muted-foreground"
			:aria-expanded="expanded"
			@click="expanded = !expanded"
		>
			<span>{{ t('pptx.textWarp.title') }}</span>
			<span class="text-muted-foreground">{{ expanded ? '-' : '+' }}</span>
		</button>

		<div v-if="!expanded" class="pptx-vue-textwarp-current text-[11px] text-muted-foreground">
			{{ currentLabel }}
		</div>

		<div v-show="expanded" class="pptx-vue-textwarp-gallery">
			<pptx-ui-ribbon-gallery
				mode="panel"
				compact
				:descriptor.prop="descriptor"
				:translateLabel.prop="translate"
				:disabled.prop="props.disabled"
				@gallery-pick="pick"
			></pptx-ui-ribbon-gallery>
		</div>
	</div>
</template>
