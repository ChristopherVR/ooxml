<script setup lang="ts">
import { buildRecordRibbon } from 'ooxml-ui/pptx';
import type { RibbonCommandRequestEvent } from 'ooxml-ui/pptx';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

const props = defineProps<{ onRecordFromBeginning: () => void; onRecordFromCurrent: () => void }>();
const { t } = useI18n();
const groups = computed(() => buildRecordRibbon(t));
function request(event: Event): void {
	const id = (event as RibbonCommandRequestEvent).detail.id;
	if (id === 'record.record.fromBeginning') {
		props.onRecordFromBeginning();
	} else if (id === 'record.record.fromCurrent') {
		props.onRecordFromCurrent();
	}
}
</script>

<template><pptx-ui-ribbon-section :groups.prop="groups" @command-request="request" /></template>
