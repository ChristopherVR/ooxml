<script setup lang="ts">
import { PowerPointViewer } from 'pptx-vue-viewer';
import { onBeforeUnmount, onMounted, ref, shallowRef } from 'vue';

import { mountOneDriveDemo } from '../../shared/onedrive-demo';
import type { OneDriveDemoDeck, OneDriveDemoMount } from '../../shared/onedrive-demo';

// `?onedrive=1`: open, edit and save a deck on your own OneDrive (docs/onedrive-quickstart.md).
const bar = ref<HTMLDivElement>();
const deck = shallowRef<OneDriveDemoDeck | null>(null);
const storage = shallowRef<OneDriveDemoMount['storage']>();
let mounted: OneDriveDemoMount | undefined;
onMounted(() => {
	mounted = mountOneDriveDemo(bar.value!, (next) => {
		deck.value = next;
	});
	storage.value = mounted.storage;
});
onBeforeUnmount(() => mounted?.dispose());
</script>

<template>
	<div ref="bar" />
	<main class="od-viewer">
		<PowerPointViewer
			v-if="deck && storage"
			:content="deck.bytes"
			:file-name="deck.fileName"
			:can-edit="deck.canEdit"
			:storage="storage"
		/>
		<p v-else class="od-empty">No OneDrive file open yet.</p>
	</main>
</template>
