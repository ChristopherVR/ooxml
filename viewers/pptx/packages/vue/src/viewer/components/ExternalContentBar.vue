<script setup lang="ts">
/**
 * ExternalContentBar: PowerPoint's security bar for a deck whose linked
 * pictures and media were not downloaded ("SECURITY WARNING ... Enable
 * Content"). A thin adapter around the shared `pptx-ui-external-content-bar`:
 * the element owns the markup and strings; this gives it the translator and
 * re-emits its two requests. Whether it shows is `barVisible` from the shared
 * `resolveExternalContentView`, decided by the viewer.
 */
import { dispatchExternalContentBarIntent } from 'ooxml-ui/pptx';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

const emit = defineEmits<{
	enable: [];
	close: [];
}>();

const { t } = useI18n();

const state = computed(() => ({ translate: t }));

function request(event: Event): void {
	dispatchExternalContentBarIntent(event, {
		enable: () => emit('enable'),
		close: () => emit('close'),
	});
}
</script>

<template>
	<pptx-ui-external-content-bar :state.prop="state" @external-content-request="request" />
</template>
