<script setup lang="ts">
import type { ThemeColorPickerCommit } from 'ooxml-ui/pptx';
import { SHAPE_LINE_SECTIONS, shapeOutlineChange } from 'ooxml-ui/pptx';
import type { PptxElement, ShapeStyle } from 'pptx-viewer-core';
import { hasShapeProperties } from 'pptx-viewer-core';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

import { injectRecentColors } from '../../composables/recent-colors-context';
import RecentColorsRow from '../RecentColorsRow.vue';
import InspectorFields from './InspectorFields.vue';
import ThemeColorSwatchGrid from './ThemeColorSwatchGrid.vue';

/**
 * StrokePanel: line/border inspector for the Vue `pptx-vue-viewer` editor.
 *
 * The line colour (picker, theme swatches, recent colours) plus the shared
 * Line options (`SHAPE_LINE_SECTIONS`: width, transparency, dash, compound,
 * join, cap) drawn by the generic {@link InspectorFields} renderer. Each
 * control emits a SHALLOW `update` patch carrying the FULL merged
 * `shapeStyle` sub-object so the parent can forward it verbatim to
 * `ops.updateElement(id, patch)`.
 *
 * Only shape-like elements (`hasShapeProperties`) expose editable controls;
 * other element types show a muted note instead.
 */
const props = defineProps<{
	element: PptxElement;
}>();

const emit = defineEmits<{
	update: [patch: Partial<PptxElement>];
}>();

const { t } = useI18n();
const recentColors = injectRecentColors();

const applicable = computed(() => hasShapeProperties(props.element));

const shapeStyle = computed<ShapeStyle | undefined>(() =>
	hasShapeProperties(props.element) ? props.element.shapeStyle : undefined,
);

const strokeColor = computed<string>(() => shapeStyle.value?.strokeColor ?? '#000000');

function patchShapeStyle(next: Partial<ShapeStyle>): void {
	emit('update', { shapeStyle: { ...shapeStyle.value, ...next } } as Partial<PptxElement>);
}

function onColor(event: Event): void {
	patchShapeStyle(
		shapeOutlineChange((event.target as HTMLInputElement).value, undefined, shapeStyle.value ?? {}),
	);
}

function onColorPick(hex: string): void {
	patchShapeStyle(shapeOutlineChange(hex, undefined, shapeStyle.value ?? {}));
	recentColors?.push(hex);
}

function onThemeColor(commit: ThemeColorPickerCommit): void {
	patchShapeStyle(shapeOutlineChange(commit.hex, commit.ref, shapeStyle.value ?? {}));
	recentColors?.push(commit.hex);
}
</script>

<template>
	<div class="pptx-vue-stroke-panel flex flex-col gap-2 p-2">
		<h3
			class="pptx-vue-stroke-title text-xs font-semibold uppercase tracking-wide text-muted-foreground"
		>
			{{ t('pptx.inspector.line') }}
		</h3>

		<p v-if="!applicable" class="pptx-vue-stroke-muted text-xs text-muted-foreground">
			{{ t('pptx.stroke.noBorderProperties') }}
		</p>

		<div v-else class="pptx-vue-stroke-fields flex flex-col gap-2">
			<label class="pptx-vue-stroke-field flex flex-col gap-1">
				<span class="pptx-vue-stroke-label text-muted-foreground">{{
					t('pptx.inspector.color')
				}}</span>
				<input
					type="color"
					class="pptx-vue-stroke-color w-full h-8 bg-muted border border-border rounded p-0.5"
					:value="strokeColor"
					@input="onColor"
					@change="onColorPick(($event.target as HTMLInputElement).value)"
				/>
			</label>
			<ThemeColorSwatchGrid
				:selected-ref="shapeStyle?.strokeColorRef"
				:selected-hex="strokeColor"
				@pick="onThemeColor"
			/>
			<RecentColorsRow
				v-if="recentColors"
				:colors="recentColors.recent.value"
				@pick="onColorPick"
			/>

			<InspectorFields
				:sections="SHAPE_LINE_SECTIONS"
				:target="shapeStyle"
				@patch="patchShapeStyle"
			/>
		</div>
	</div>
</template>
