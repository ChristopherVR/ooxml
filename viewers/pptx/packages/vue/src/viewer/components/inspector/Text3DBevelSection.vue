<script setup lang="ts">
import { BEVEL_PRESETS, TEXT_3D_MAX_BEVEL_PT, clampText3dPt, text3dEmuToPt } from 'ooxml-ui/pptx';
import type { BevelPresetType } from 'pptx-viewer-core';
import { useI18n } from 'vue-i18n';

/**
 * Text3DBevelSection: one bevel editor (type / width / height), reused for the
 * top and bottom bevels of `Text3DProperties`. Widths and heights are shown in
 * points and emitted in points (clamped by the shared `text-3d-fields` limits);
 * the parent converts to EMU. Mirrors the React `BevelSection` sub-component.
 */
const props = defineProps<{
	label: string;
	bevelType: BevelPresetType | undefined;
	bevelWidth: number | undefined;
	bevelHeight: number | undefined;
	disabled?: boolean;
}>();

const emit = defineEmits<{
	typeChange: [value: BevelPresetType];
	widthChange: [value: number];
	heightChange: [value: number];
}>();

const { t } = useI18n();

const INPUT_CLS = 'bg-muted border border-border rounded px-2 py-1 max-md:min-h-[44px]';

function onType(event: Event): void {
	emit('typeChange', (event.target as HTMLSelectElement).value as BevelPresetType);
}
function onWidth(event: Event): void {
	const v = Number((event.target as HTMLInputElement).value);
	if (Number.isFinite(v)) {
		emit('widthChange', clampText3dPt(v, TEXT_3D_MAX_BEVEL_PT));
	}
}
function onHeight(event: Event): void {
	const v = Number((event.target as HTMLInputElement).value);
	if (Number.isFinite(v)) {
		emit('heightChange', clampText3dPt(v, TEXT_3D_MAX_BEVEL_PT));
	}
}
</script>

<template>
	<div class="pptx-vue-text3d-bevel space-y-1 pl-4">
		<span class="text-[11px] text-muted-foreground">{{ props.label }}</span>
		<div class="grid grid-cols-3 gap-2">
			<label class="flex flex-col gap-1">
				<span class="text-muted-foreground">{{ t('pptx.text3d.bevelType') }}</span>
				<pptx-ui-select
					:aria-label="t('pptx.text3d.bevelType')"
					:class="INPUT_CLS"
					:value="props.bevelType ?? 'none'"
					:disabled="props.disabled"
					@change="onType"
				>
					<option v-for="opt in BEVEL_PRESETS" :key="opt.value" :value="opt.value">
						{{ t(opt.labelKey) }}
					</option>
				</pptx-ui-select>
			</label>
			<label class="flex flex-col gap-1">
				<span class="text-muted-foreground">{{ t('pptx.text3d.bevelWidth') }}</span>
				<input
					type="number"
					min="0"
					:max="TEXT_3D_MAX_BEVEL_PT"
					step="1"
					:class="INPUT_CLS"
					:disabled="props.disabled"
					:value="text3dEmuToPt(props.bevelWidth)"
					@input="onWidth"
				/>
			</label>
			<label class="flex flex-col gap-1">
				<span class="text-muted-foreground">{{ t('pptx.text3d.bevelHeight') }}</span>
				<input
					type="number"
					min="0"
					:max="TEXT_3D_MAX_BEVEL_PT"
					step="1"
					:class="INPUT_CLS"
					:disabled="props.disabled"
					:value="text3dEmuToPt(props.bevelHeight)"
					@input="onHeight"
				/>
			</label>
		</div>
	</div>
</template>
