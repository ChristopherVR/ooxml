<script setup lang="ts">
import { Lock, LockOpen } from 'lucide-vue-next';
import { ELEMENT_GEOMETRY_SECTIONS, elementLockTogglePatch, isElementLocked } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

import InspectorFields from './InspectorFields.vue';

/**
 * ArrangePanel: position, size, rotation, flip controls and element lock toggle.
 * X / Y / W / H / rotation are the shared `ELEMENT_GEOMETRY_SECTIONS`
 * descriptors drawn by the generic {@link InspectorFields} renderer.
 *
 * Applicable to every element (all carry `x`/`y`/`width`/`height`/`rotation`
 * and the flip flags). Emits SHALLOW patches that the parent merges onto the
 * element via `ops.updateElement(id, patch)`.
 */
const props = defineProps<{
	element: PptxElement;
	canEdit?: boolean;
}>();

const emit = defineEmits<{
	update: [patch: Partial<PptxElement>];
}>();

const { t } = useI18n();

const flipHorizontal = computed(() => Boolean(props.element.flipHorizontal));
const flipVertical = computed(() => Boolean(props.element.flipVertical));

// Shared decides what counts as "locked" (`noSelect` / `noMove` / `noResize`),
// so the toggle's state cannot drift from what the canvas actually enforces.
const isLocked = computed(() => isElementLocked(props.element));

function toggleLock(): void {
	if (!props.canEdit) {
		return;
	}
	emit('update', { locks: elementLockTogglePatch(!isLocked.value) } as Partial<PptxElement>);
}

function onFlipHorizontal(checked: boolean): void {
	emit('update', { flipHorizontal: checked });
}

function onFlipVertical(checked: boolean): void {
	emit('update', { flipVertical: checked });
}
</script>

<template>
	<div class="pptx-vue-arrange flex flex-col gap-2 text-xs">
		<div class="flex items-center justify-between">
			<span class="text-muted-foreground font-medium">{{ t('pptx.arrange.positionSize') }}</span>
			<button
				type="button"
				:disabled="!props.canEdit"
				:title="isLocked ? t('pptx.arrange.unlockElement') : t('pptx.arrange.lockElement')"
				:aria-label="isLocked ? t('pptx.arrange.unlockElement') : t('pptx.arrange.lockElement')"
				:aria-pressed="isLocked"
				class="p-1 rounded hover:bg-accent disabled:opacity-40 disabled:cursor-not-allowed transition-colors flex items-center justify-center max-md:min-w-[44px]! max-md:min-h-[44px]!"
				@click="toggleLock"
			>
				<Lock v-if="isLocked" class="w-3.5 h-3.5 text-amber-400" />
				<LockOpen v-else class="w-3.5 h-3.5 text-muted-foreground" />
			</button>
		</div>
		<InspectorFields
			:sections="ELEMENT_GEOMETRY_SECTIONS"
			:target="element"
			:disabled="!props.canEdit"
			@patch="(patch) => emit('update', patch as Partial<PptxElement>)"
		/>

		<div class="pptx-vue-arrange-flips grid grid-cols-2 gap-2">
			<label class="pptx-vue-arrange-check inline-flex items-center gap-2 text-foreground">
				<pptx-ui-checkbox
					:checked="flipHorizontal"
					@change="onFlipHorizontal(($event.target as HTMLInputElement).checked)"
				/>
				{{ t('pptx.arrange.flipHorizontally') }}
			</label>
			<label class="pptx-vue-arrange-check inline-flex items-center gap-2 text-foreground">
				<pptx-ui-checkbox
					:checked="flipVertical"
					@change="onFlipVertical(($event.target as HTMLInputElement).checked)"
				/>
				{{ t('pptx.arrange.flipVertically') }}
			</label>
		</div>
	</div>
</template>
