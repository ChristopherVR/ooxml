<script setup lang="ts">
import {
	fillColorOf,
	resolveShapeFillView,
	SHAPE_FILL_MODE_SECTIONS,
	SHAPE_FILL_OPTIONS_SECTIONS,
	shapeFillChange,
	shapeFillPatch,
	shapeStylePatch,
} from 'ooxml-ui/pptx';
import type { ShapeFillAction, ThemeColorPickerCommit } from 'ooxml-ui/pptx';
import type { PptxElement, ShapeStyle } from 'pptx-viewer-core';
import { hasShapeProperties } from 'pptx-viewer-core';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

import { useLinkedContentUrl } from '../../composables/linked-content';
import { injectRecentColors } from '../../composables/recent-colors-context';
import { useLiveElementColor } from '../../composables/useLiveElementColor';
import RecentColorsRow from '../RecentColorsRow.vue';
import FillGradientControls from './FillGradientControls.vue';
import InspectorFields from './InspectorFields.vue';
import ThemeColorSwatchGrid from './ThemeColorSwatchGrid.vue';

/**
 * FillPanel: the Format pane's Fill section. The fill kind, transparency,
 * gradient type and angle and picture tiling are the shared fill descriptors
 * drawn by {@link InspectorFields}; the solid colour, the gradient stops
 * ({@link FillGradientControls}), the pattern swatches and colours and the
 * picture source are what `resolveShapeFillView` asks for, and their events
 * are the shared `shapeFillPatch` actions. Every patch is emitted as the full
 * merged `shapeStyle` (`shapeStylePatch`). A muted note stands in for
 * elements without shape properties.
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

const style = computed<ShapeStyle | undefined>(() =>
	hasShapeProperties(props.element) ? props.element.shapeStyle : undefined,
);

const view = computed(() => resolveShapeFillView(style.value));
/** The picture fill preview: a linked picture only once this viewer allows linked content. */
const imagePreview = useLinkedContentUrl(() => view.value.image?.url);

function patchStyle(patch: Partial<ShapeStyle>): void {
	emit('update', shapeStylePatch(props.element, patch));
}

function act(action: ShapeFillAction): void {
	const patch = shapeFillPatch(style.value, action);
	if (patch) {
		patchStyle(patch);
	}
}

/** The native picker: live while dragging, one `update` and one recent colour per pick. */
const fillColorEvents = useLiveElementColor({
	element: () => props.element,
	colorOf: fillColorOf,
	patch: (el, hex) => shapeStylePatch(el, shapeFillChange(hex)),
	commit: (patch) => emit('update', patch),
});

function onThemeColor(commit: ThemeColorPickerCommit): void {
	act({ type: 'color', hex: commit.hex, ref: commit.ref });
	recentColors?.push(commit.hex);
}

function onRecentColor(hex: string): void {
	act({ type: 'color', hex });
	recentColors?.push(hex);
}

function onImageFile(event: Event): void {
	const file = (event.target as HTMLInputElement).files?.[0];
	if (!file) {
		return;
	}
	const reader = new FileReader();
	reader.onload = () => act({ type: 'imageUrl', url: String(reader.result) });
	reader.readAsDataURL(file);
}

const valueOf = (event: Event): string => (event.target as HTMLInputElement).value;
</script>

<template>
	<div class="pptx-vue-fill flex flex-col gap-2 text-xs">
		<p v-if="!applicable" class="pptx-vue-fill-note text-muted-foreground italic">
			{{ t('pptx.fill.noOptions') }}
		</p>

		<template v-else>
			<InspectorFields :sections="SHAPE_FILL_MODE_SECTIONS" :target="style" @patch="patchStyle" />

			<template v-if="view.color">
				<label class="pptx-vue-fill-field flex flex-col gap-1">
					<span class="pptx-vue-fill-label text-muted-foreground">{{ t('pptx.fill.color') }}</span>
					<input
						type="color"
						class="pptx-vue-fill-color w-full h-8 p-0 bg-muted border border-border rounded"
						:value="view.color.value"
						v-on="fillColorEvents"
					/>
				</label>
				<ThemeColorSwatchGrid
					:selected-ref="view.color.colorRef"
					:selected-hex="view.color.value"
					@pick="onThemeColor"
				/>
				<RecentColorsRow
					v-if="recentColors"
					:colors="recentColors.recent.value"
					@pick="onRecentColor"
				/>
			</template>

			<InspectorFields
				:sections="SHAPE_FILL_OPTIONS_SECTIONS"
				:target="style"
				@patch="patchStyle"
			/>

			<FillGradientControls v-if="view.stops" :stops="view.stops" @action="act" />

			<template v-if="view.pattern">
				<span class="pptx-vue-pattern-label text-muted-foreground">{{
					t('pptx.table.patternPreset')
				}}</span>
				<div
					class="pptx-vue-pattern-grid grid max-h-48 grid-cols-8 gap-1 overflow-y-auto rounded border border-border p-2"
				>
					<button
						v-for="swatch in view.pattern.swatches"
						:key="swatch.value"
						type="button"
						data-testid="fill-pattern-swatch"
						:title="t(swatch.labelKey)"
						:aria-label="t(swatch.labelKey)"
						:aria-pressed="swatch.selected"
						class="pptx-vue-pattern-swatch h-8 w-8 rounded border transition-colors"
						:class="
							swatch.selected ? 'border-2 border-primary' : 'border-border hover:border-primary/50'
						"
						@click="act({ type: 'patternPreset', preset: swatch.value })"
					>
						<div
							v-if="swatch.backgroundImage"
							class="h-full w-full rounded-sm"
							:style="{
								backgroundImage: swatch.backgroundImage,
								backgroundRepeat: 'repeat',
								backgroundSize: '8px 8px',
							}"
						/>
					</button>
				</div>
				<label class="pptx-vue-pattern-field flex flex-col gap-1">
					<span class="pptx-vue-pattern-label text-muted-foreground">{{
						t('pptx.fillAdvanced.foregroundColor')
					}}</span>
					<input
						type="color"
						class="pptx-vue-pattern-color h-8 w-full rounded border border-border bg-muted p-0"
						:value="view.pattern.foreground"
						@input="act({ type: 'patternForeground', hex: valueOf($event) })"
						@change="recentColors?.push(valueOf($event))"
					/>
				</label>
				<label class="pptx-vue-pattern-field flex flex-col gap-1">
					<span class="pptx-vue-pattern-label text-muted-foreground">{{
						t('pptx.fillAdvanced.backgroundColor')
					}}</span>
					<input
						type="color"
						class="pptx-vue-pattern-color h-8 w-full rounded border border-border bg-muted p-0"
						:value="view.pattern.background"
						@input="act({ type: 'patternBackground', hex: valueOf($event) })"
						@change="recentColors?.push(valueOf($event))"
					/>
				</label>
			</template>

			<template v-if="view.image">
				<label class="pptx-vue-fill-field flex flex-col gap-1">
					<span class="pptx-vue-fill-label text-muted-foreground">{{
						t('pptx.fillAdvanced.imageUrl')
					}}</span>
					<input
						type="text"
						data-testid="fill-image-url"
						class="rounded border border-border bg-muted px-2 py-1"
						placeholder="https://example.com/image.png"
						:value="view.image.url"
						@change="act({ type: 'imageUrl', url: valueOf($event) })"
					/>
				</label>
				<label class="pptx-vue-fill-field flex flex-col gap-1">
					<span class="pptx-vue-fill-label text-muted-foreground">{{
						t('pptx.fillAdvanced.orChooseFile')
					}}</span>
					<input type="file" accept="image/*" class="text-xs" @change="onImageFile" />
				</label>
				<img
					v-if="imagePreview"
					:src="imagePreview"
					:alt="t('pptx.fillAdvanced.fillPreviewAlt')"
					class="h-16 w-full rounded border border-border object-cover"
				/>
			</template>
		</template>
	</div>
</template>
