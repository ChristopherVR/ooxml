<script setup lang="ts">
import type { ShapeFillAction, ShapeFillStopView, ThemeColorPickerCommit } from 'ooxml-ui/pptx';
import { useI18n } from 'vue-i18n';

import { injectRecentColors } from '../../composables/recent-colors-context';
import ThemeColorSwatchGrid from './ThemeColorSwatchGrid.vue';

/**
 * FillGradientControls: the gradient stop list of {@link FillPanel}, one row
 * per shared `ShapeFillStopView` (colour with its theme grid, position,
 * opacity, remove) and an Add Stop button. Every control emits a shared
 * `ShapeFillAction`; FillPanel turns it into the style patch.
 */
defineProps<{
	stops: readonly ShapeFillStopView[];
}>();

const emit = defineEmits<{
	action: [action: ShapeFillAction];
}>();

const { t } = useI18n();
const recentColors = injectRecentColors();

const valueOf = (event: Event): string => (event.target as HTMLInputElement).value;

function onThemePick(index: number, commit: ThemeColorPickerCommit): void {
	emit('action', { type: 'stopColor', index, hex: commit.hex, ref: commit.ref });
	recentColors?.push(commit.hex);
}
</script>

<template>
	<div class="pptx-vue-gradient flex flex-col gap-2 text-xs">
		<span class="pptx-vue-gradient-stops-heading uppercase tracking-wide text-muted-foreground">
			{{ t('pptx.fillAdvanced.gradientStops') }}
		</span>
		<div v-for="stop in stops" :key="stop.index" class="pptx-vue-gradient-stop-block">
			<div
				data-testid="fill-gradient-stop"
				class="pptx-vue-gradient-stop-row flex items-center gap-1"
			>
				<span class="pptx-vue-gradient-stop-idx w-3 text-center text-muted-foreground">{{
					stop.index + 1
				}}</span>
				<input
					type="color"
					class="pptx-vue-gradient-color h-6 w-8 cursor-pointer rounded border border-border"
					:aria-label="t('pptx.gradient.stops')"
					:value="stop.color"
					@input="emit('action', { type: 'stopColor', index: stop.index, hex: valueOf($event) })"
					@change="recentColors?.push(valueOf($event))"
				/>
				<input
					type="number"
					class="pptx-vue-gradient-input w-14 rounded border border-border bg-muted px-1.5 py-0.5"
					:aria-label="t('pptx.gradient.position')"
					min="0"
					max="100"
					:value="stop.position"
					@change="
						emit('action', {
							type: 'stopPosition',
							index: stop.index,
							position: Number(valueOf($event)),
						})
					"
				/>
				<span class="text-[10px] text-muted-foreground">%</span>
				<input
					type="number"
					class="pptx-vue-gradient-input w-14 rounded border border-border bg-muted px-1.5 py-0.5"
					:aria-label="t('pptx.strokeEffects.fillOpacity')"
					min="0"
					max="100"
					:value="stop.opacity"
					@change="
						emit('action', {
							type: 'stopOpacity',
							index: stop.index,
							opacity: Number(valueOf($event)),
						})
					"
				/>
				<button
					type="button"
					class="pptx-vue-gradient-remove ml-auto rounded border border-border px-1.5 text-muted-foreground hover:text-destructive"
					:title="t('pptx.gradient.removeStop')"
					:aria-label="t('pptx.gradient.removeStop')"
					:disabled="!stop.removable"
					@click="emit('action', { type: 'removeStop', index: stop.index })"
				>
					&times;
				</button>
			</div>
			<ThemeColorSwatchGrid
				:selected-ref="stop.colorRef"
				:selected-hex="stop.color"
				@pick="(commit) => onThemePick(stop.index, commit)"
			/>
		</div>
		<button
			type="button"
			class="pptx-vue-gradient-add self-start rounded border border-border px-2 py-1 text-[11px] text-primary hover:bg-accent"
			@click="emit('action', { type: 'addStop' })"
		>
			{{ t('pptx.gradient.addStop') }}
		</button>
	</div>
</template>
