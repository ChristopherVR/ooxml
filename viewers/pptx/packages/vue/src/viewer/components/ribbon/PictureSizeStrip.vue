<script setup lang="ts">
/**
 * PictureSizeStrip: Picture Format > Size, the shared
 * `pptx-ui-ribbon-home-picture-size` strip. Its Crop split button drives the
 * same injected crop controller as Home > Arrange.
 */
import type { RibbonHomeRequestEvent, ToolbarActionId } from 'ooxml-ui/pptx';
import { homeSnapshotTranslator, pictureSizeHomeControls } from 'ooxml-ui/pptx';
import { computed, inject } from 'vue';
import { useI18n } from 'vue-i18n';

import { runCropRequest } from '../../composables/crop-request';
import { MergeCropKey } from '../../composables/merge-crop-context';
import { useToolbarVisibility } from '../../composables/useToolbarVisibility';

const props = defineProps<{ hiddenActions?: readonly ToolbarActionId[] }>();
const { t, locale } = useI18n();
const { isHidden } = useToolbarVisibility(() => props.hiddenActions);
const controller = inject(MergeCropKey, undefined);

const state = computed(() => ({
	controls: pictureSizeHomeControls({
		editable: Boolean(controller),
		canCrop: Boolean(controller?.canCrop.value),
		cropActive: Boolean(controller?.cropActive.value),
		hideCrop: isHidden('crop'),
	}),
	// The locale is read so a language switch re-translates the shared labels.
	locale: locale.value,
	translate: homeSnapshotTranslator(['picture-size'], t),
}));

function request(event: RibbonHomeRequestEvent): void {
	runCropRequest(controller, event.detail.value);
}
</script>

<template>
	<!-- A press here must not count as "outside" and commit crop mode. -->
	<div class="contents" data-pptx-crop-keep="true">
		<pptx-ui-ribbon-home-picture-size :state.prop="state" @home-request="request" />
	</div>
</template>
