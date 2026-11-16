<script setup lang="ts">
import { ELEMENT_ALT_TEXT_SECTIONS, getNonVisualDescriptionFields } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { computed } from 'vue';

import InspectorFields from './InspectorFields.vue';

/**
 * AccessibilityPanel: the Alt Text fields of a shape, text box, connector or
 * graphic frame, the shared `ELEMENT_ALT_TEXT_SECTIONS` drawn by the generic
 * renderer. A picture's alt text is drawn in `ImagePanel` from the same
 * descriptors.
 */
const props = defineProps<{
	element: PptxElement;
	canEdit?: boolean;
}>();

const emit = defineEmits<{
	update: [patch: Partial<PptxElement>];
}>();

const fields = computed(() => getNonVisualDescriptionFields(props.element));
</script>

<template>
	<div
		v-if="fields.showAltText || fields.showTitle"
		class="pptx-vue-accessibility-panel flex flex-col gap-2 text-[11px]"
		data-pptx-accessibility-text
	>
		<InspectorFields
			:sections="ELEMENT_ALT_TEXT_SECTIONS"
			:target="props.element"
			:disabled="!props.canEdit"
			@patch="emit('update', $event)"
		/>
	</div>
</template>
