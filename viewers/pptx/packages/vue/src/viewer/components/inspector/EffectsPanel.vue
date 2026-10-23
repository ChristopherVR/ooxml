<script setup lang="ts">
import { SHAPE_EFFECT_SECTIONS, shapeQuickStylePatch } from 'ooxml-ui/pptx';
import type { PptxElement, ShapeStyle } from 'pptx-viewer-core';
import { hasShapeProperties } from 'pptx-viewer-core';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

import { injectInspectorPreview } from '../../composables/inspector-preview-context';
import InspectorFields from './InspectorFields.vue';
import QuickStylesGallery from './QuickStylesGallery.vue';

/**
 * EffectsPanel: element opacity plus every shape visual effect PowerPoint's
 * Format Shape > Effects pane exposes (outer and inner shadow, glow, soft
 * edge, reflection, blur, top bevel). The effect fields are the shared
 * `SHAPE_EFFECT_SECTIONS` descriptors drawn by the generic
 * {@link InspectorFields} renderer; their `Partial<ShapeStyle>` patches are
 * merged into the element with the shared `shapeQuickStylePatch`. This
 * component itself only owns opacity (the one field that is NOT part of
 * `shapeStyle`) and the shape-like gate.
 *
 * Shadow/glow/reflection/soft-edge controls only apply to shape-like elements
 * (core `hasShapeProperties`). For other elements a muted note is shown.
 */
const props = withDefaults(
	defineProps<{
		element: PptxElement;
		/** Read-only decks keep the Quick Styles tiles visible but unpickable. */
		canEdit?: boolean;
	}>(),
	{ canEdit: true },
);

const emit = defineEmits<{
	update: [patch: Partial<PptxElement>];
}>();

const { t } = useI18n();

const isShapeLike = computed(() => hasShapeProperties(props.element));
const shapeStyle = computed(() =>
	hasShapeProperties(props.element) ? props.element.shapeStyle : undefined,
);

/** Merge an effect field's shape-style patch into the element. */
function onEffectPatch(patch: Partial<ShapeStyle>): void {
	emit('update', shapeQuickStylePatch(props.element, patch));
}

const preview = injectInspectorPreview();

/** The shape style of element `id` now: what a drag the selection left is put back on. */
function styleOf(id: string): ShapeStyle | undefined {
	const owner = preview?.element(id);
	return owner && hasShapeProperties(owner) ? owner.shapeStyle : undefined;
}

/**
 * An effect colour drag's live shade, merged into the style of the element it is
 * for (the current one, or one the drag has left): on the slide at once, no undo step.
 */
function onEffectPreview(patch: Partial<ShapeStyle>, target: string): void {
	const owner = preview?.element(target);
	if (owner) {
		preview?.preview(shapeQuickStylePatch(owner, patch), target);
	}
}

/**
 * Apply a Quick Styles gallery preset (a `Partial<ShapeStyle>` of unrelated
 * fields, e.g. fill/outline) on top of the CURRENT shapeStyle, through the
 * shared `shapeQuickStylePatch` merge.
 */
function onQuickStyleSelect(preset: Partial<ShapeStyle>): void {
	emit('update', shapeQuickStylePatch(props.element, preset));
}

// ---------------------------------------------------------------------------
// Opacity (element-level, shallow patch)
// ---------------------------------------------------------------------------

const opacityPercent = computed(() => Math.round((props.element.opacity ?? 1) * 100));

function onOpacity(value: string): void {
	const n = Number(value);
	if (!Number.isFinite(n)) {
		return;
	}
	emit('update', { opacity: Math.max(0, Math.min(100, n)) / 100 });
}
</script>

<template>
	<div class="pptx-vue-effects flex flex-col gap-3 text-xs">
		<label class="pptx-vue-effects-field flex flex-col gap-1">
			<span class="pptx-vue-effects-label text-muted-foreground">{{
				t('pptx.effects.opacityPercent', { value: opacityPercent })
			}}</span>
			<input
				type="range"
				class="pptx-vue-effects-range w-full accent-primary"
				min="0"
				max="100"
				:value="opacityPercent"
				@input="onOpacity(($event.target as HTMLInputElement).value)"
			/>
		</label>

		<template v-if="isShapeLike">
			<QuickStylesGallery :disabled="!props.canEdit" @select="onQuickStyleSelect" />
			<InspectorFields
				:sections="SHAPE_EFFECT_SECTIONS"
				:target="shapeStyle"
				:target-key="props.element.id"
				:lookup="styleOf"
				@patch="onEffectPatch"
				@preview="onEffectPreview"
			/>
		</template>

		<p v-else class="pptx-vue-effects-note text-muted-foreground italic">
			{{ t('pptx.effects.shapeOnlyNote') }}
		</p>
	</div>
</template>
