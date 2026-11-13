<script setup lang="ts">
import {
	elementRenderErrorContext,
	elementRenderRetryKey,
	reportElementRenderError,
	shouldRetryElementRender,
} from 'ooxml-ui/pptx';
import type { ElementRenderRetryKey } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { computed, onErrorCaptured, shallowRef } from 'vue';

import {
	injectElementRenderErrorScope,
	isEventHandlerErrorInfo,
} from '../composables/element-render-errors';

/**
 * ElementBoundary: omits one slide element whose render, setup or lifecycle hook throws, so
 * its neighbours, the slide and the viewer stay. The host hears about it through the injected
 * `onElementRenderError` (default `console.error`). A replacement element object or a changed
 * id renders it again; re-rendering the same failed object does not. Event handler errors are
 * not caught here (they are not render failures), matching React's boundaries.
 */
const props = defineProps<{ element: PptxElement }>();

const scope = injectElementRenderErrorScope();
const failed = shallowRef<ElementRenderRetryKey | null>(null);
const omitted = computed(() => !shouldRetryElementRender(failed.value, props.element));

onErrorCaptured((error, _instance, info) => {
	if (isEventHandlerErrorInfo(info)) {
		return undefined;
	}
	failed.value = elementRenderRetryKey(props.element);
	const { slideId, onElementRenderError } = scope();
	reportElementRenderError(
		onElementRenderError,
		error,
		elementRenderErrorContext(props.element, slideId),
	);
	return false;
});
</script>

<template>
	<slot v-if="!omitted" />
</template>
