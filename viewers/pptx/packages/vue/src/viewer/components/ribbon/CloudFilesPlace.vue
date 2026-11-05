<script setup lang="ts">
/** One File > Open place for a cloud provider: the shared `pptx-ui-cloud-files` browser. */
import { mountCloudFilesPlace } from 'ooxml-ui/pptx';
import type {
	CloudFileProvider,
	CloudFilesPlace,
	CloudFilesPlaceOptions,
	CloudOpenedFile,
} from 'ooxml-ui/pptx';
import { onBeforeUnmount, onMounted, ref, watch } from 'vue';
import { useI18n } from 'vue-i18n';

const props = defineProps<{
	provider: CloudFileProvider;
	onOpen: (file: CloudOpenedFile) => void;
}>();

const { t, locale } = useI18n();
const host = ref<HTMLDivElement | null>(null);
let place: CloudFilesPlace | null = null;

function options(): CloudFilesPlaceOptions {
	return {
		provider: props.provider,
		translate: (key, params) => (params ? t(key, params) : t(key)),
		locale: locale.value,
		onOpen: (file) => props.onOpen(file),
	};
}

function mountPlace(): void {
	place?.destroy();
	place = host.value ? mountCloudFilesPlace(host.value, options()) : null;
}

onMounted(mountPlace);
// A new provider gets a fresh element; a language change re-labels the current one.
watch(() => props.provider, mountPlace);
watch(locale, () => place?.update(options()));
onBeforeUnmount(() => {
	place?.destroy();
	place = null;
});
</script>

<template>
	<div ref="host" :data-pptx-cloud-place="props.provider.id" class="mt-6 max-w-[900px]" />
</template>
