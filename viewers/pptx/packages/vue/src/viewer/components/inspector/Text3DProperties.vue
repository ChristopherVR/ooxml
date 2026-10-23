<script setup lang="ts">
import {
	MATERIAL_PRESETS,
	TEXT_3D_MAX_EXTRUSION_PT,
	clampText3dPt,
	hasText3dExtrusion,
	mergeText3d,
	normalizeHexColor,
	text3dEmuToPt,
	text3dFieldsPatch,
	text3dPtToEmu,
	toggleText3dExtrusion,
} from 'ooxml-ui/pptx';
import type { BevelPresetType, MaterialPresetType, Text3DStyle, TextStyle } from 'pptx-viewer-core';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

import { injectRecentColors } from '../../composables/recent-colors-context';
import Text3DBevelSection from './Text3DBevelSection.vue';

/**
 * Text3DProperties: 3D-text (WordArt) authoring UI, mirroring the React
 * `Text3DProperties`. Exposes an extrusion toggle with depth/colour, top and
 * bottom bevels, and a surface material. The EMU<->pt conversions, limits and
 * the seeded depth are the shared `text-3d-fields` helpers. Emits `update` with
 * a PARTIAL `TextStyle` patch carrying the FULL merged `text3d` sub-object and
 * its `flatText` marker; the parent (`TextPanel`) merges into `textStyle`.
 */
const props = withDefaults(
	defineProps<{
		ts: TextStyle | undefined;
		/**
		 * The shape inherits `a:sp3d` from its layout/master (a placeholder), so
		 * switching 3-D off has to write `a:flatTx` (`text3dInheritsFromTemplate`).
		 */
		inherits?: boolean;
		/** Read-only decks show the 3-D settings without letting them change. */
		disabled?: boolean;
	}>(),
	{ inherits: false, disabled: false },
);

const emit = defineEmits<{
	update: [patch: Partial<TextStyle>];
}>();

const { t } = useI18n();
const recentColors = injectRecentColors();

const t3d = computed<Text3DStyle | undefined>(() => props.ts?.text3d);
const hasExtrusion = computed(() => hasText3dExtrusion(t3d.value));

function emitText3d(next: Text3DStyle | undefined): void {
	emit('update', text3dFieldsPatch(next, props.inherits));
}

function update3d(partial: Partial<Text3DStyle>): void {
	emitText3d(mergeText3d(t3d.value, partial));
}

function toggleExtrusion(checked: boolean): void {
	emitText3d(toggleText3dExtrusion(t3d.value, checked));
}

function onDepth(event: Event): void {
	const v = Number((event.target as HTMLInputElement).value);
	if (Number.isFinite(v)) {
		update3d({ extrusionHeight: text3dPtToEmu(clampText3dPt(v, TEXT_3D_MAX_EXTRUSION_PT)) });
	}
}

function onExtrusionColorCommit(event: Event): void {
	recentColors?.push((event.target as HTMLInputElement).value);
}

function onMaterial(event: Event): void {
	const v = (event.target as HTMLSelectElement).value;
	update3d({ presetMaterial: v ? (v as MaterialPresetType) : undefined });
}

const INPUT_CLS = 'bg-muted border border-border rounded px-2 py-1 max-md:min-h-[44px]';
const COLOR_CLS = 'h-8 bg-muted border border-border rounded px-1';
</script>

<template>
	<div class="pptx-vue-text3d mt-2 rounded border border-border bg-card p-2 space-y-2">
		<div class="text-[11px] uppercase tracking-wide text-muted-foreground">
			{{ t('pptx.text3d.title') }}
		</div>

		<!-- Extrusion toggle -->
		<div class="space-y-1.5">
			<label class="inline-flex items-center gap-2 text-foreground">
				<pptx-ui-checkbox
					:checked="hasExtrusion"
					:disabled="props.disabled"
					@change="toggleExtrusion(($event.target as HTMLInputElement).checked)"
				/>
				{{ t('pptx.text3d.extrusion') }}
			</label>
			<div v-if="hasExtrusion" class="grid grid-cols-2 gap-2 pl-4">
				<label class="flex flex-col gap-1">
					<span class="text-muted-foreground">{{ t('pptx.text3d.extrusionDepth') }}</span>
					<input
						type="number"
						min="0"
						:max="TEXT_3D_MAX_EXTRUSION_PT"
						step="1"
						:class="INPUT_CLS"
						:aria-label="t('pptx.text3d.extrusionDepth')"
						:disabled="props.disabled"
						:value="text3dEmuToPt(t3d?.extrusionHeight)"
						@input="onDepth"
					/>
				</label>
				<label class="flex flex-col gap-1">
					<span class="text-muted-foreground">{{ t('pptx.text3d.extrusionColor') }}</span>
					<input
						type="color"
						:class="COLOR_CLS"
						:aria-label="t('pptx.text3d.extrusionColor')"
						:disabled="props.disabled"
						:value="normalizeHexColor(t3d?.extrusionColor, '#888888')"
						@input="update3d({ extrusionColor: ($event.target as HTMLInputElement).value })"
						@change="onExtrusionColorCommit"
					/>
				</label>
			</div>
		</div>

		<template v-if="hasExtrusion">
			<Text3DBevelSection
				:label="t('pptx.text3d.bevelTop')"
				:bevel-type="t3d?.bevelTopType"
				:bevel-width="t3d?.bevelTopWidth"
				:bevel-height="t3d?.bevelTopHeight"
				:disabled="props.disabled"
				@type-change="(v: BevelPresetType) => update3d({ bevelTopType: v || undefined })"
				@width-change="(v: number) => update3d({ bevelTopWidth: text3dPtToEmu(v) })"
				@height-change="(v: number) => update3d({ bevelTopHeight: text3dPtToEmu(v) })"
			/>
			<Text3DBevelSection
				:label="t('pptx.text3d.bevelBottom')"
				:bevel-type="t3d?.bevelBottomType"
				:bevel-width="t3d?.bevelBottomWidth"
				:bevel-height="t3d?.bevelBottomHeight"
				:disabled="props.disabled"
				@type-change="(v: BevelPresetType) => update3d({ bevelBottomType: v || undefined })"
				@width-change="(v: number) => update3d({ bevelBottomWidth: text3dPtToEmu(v) })"
				@height-change="(v: number) => update3d({ bevelBottomHeight: text3dPtToEmu(v) })"
			/>

			<label class="flex flex-col gap-1 pl-4">
				<span class="text-muted-foreground">{{ t('pptx.text3d.material') }}</span>
				<pptx-ui-select
					:aria-label="t('pptx.text3d.material')"
					:class="INPUT_CLS"
					:value="t3d?.presetMaterial ?? ''"
					:disabled="props.disabled"
					@change="onMaterial"
				>
					<option v-for="opt in MATERIAL_PRESETS" :key="opt.value" :value="opt.value">
						{{ t(opt.labelKey) }}
					</option>
				</pptx-ui-select>
			</label>
		</template>
	</div>
</template>
