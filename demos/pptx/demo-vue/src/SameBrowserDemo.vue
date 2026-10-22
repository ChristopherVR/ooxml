<script setup lang="ts">
import { PowerPointViewer } from 'pptx-vue-viewer';
import { onBeforeUnmount, onMounted, shallowRef } from 'vue';

import { createSameBrowserDemo } from '../../shared/same-browser-collaboration';
import type { SameBrowserDemo } from '../../shared/same-browser-collaboration';

const demo = shallowRef<SameBrowserDemo | null>(null);
const error = shallowRef('');
let disposed = false;
onMounted(() => {
	void createSameBrowserDemo()
		.then((value) => {
			if (disposed) {
				value.dispose();
				return;
			}
			demo.value = value;
			return undefined;
		})
		.catch((reason: unknown) => {
			error.value = String(reason);
		});
});
onBeforeUnmount(() => {
	disposed = true;
	demo.value?.dispose();
});
</script>

<template>
	<main :style="{ position: 'fixed', inset: '28px 0 0' }">
		<p v-if="error" role="alert">{{ error }}</p>
		<PowerPointViewer
			v-if="demo"
			:content="demo.source"
			:file-name="demo.fileName"
			:collaboration="demo.collaboration"
			:can-edit="true"
			@collaboration-change="demo.report"
		/>
	</main>
</template>
