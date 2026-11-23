<script setup lang="ts">
import { buildRecordRibbon, runRecordRibbonAction } from 'ooxml-ui/pptx';
import type { RecordTarget, RibbonCommandRequestEvent } from 'ooxml-ui/pptx';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

const props = defineProps<{
	onRecordFromBeginning: () => void;
	onRecordFromCurrent: () => void;
	getTarget?: (() => RecordTarget | undefined) | undefined;
}>();
const { t } = useI18n();
const groups = computed(() => buildRecordRibbon(t, props.getTarget?.()));
function request(event: Event): void {
	const id = (event as RibbonCommandRequestEvent).detail.id;
	if (id === 'record.record.fromBeginning') {
		props.onRecordFromBeginning();
	} else if (id === 'record.record.fromCurrent') {
		props.onRecordFromCurrent();
	} else if (event.currentTarget instanceof HTMLElement)
		runRecordRibbonAction(event.currentTarget, id, () => props.getTarget?.(), t);
}
</script>

<template><pptx-ui-ribbon-section :groups.prop="groups" @command-request="request" /></template>
