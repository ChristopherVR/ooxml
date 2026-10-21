<script setup lang="ts">
import { buildHelpRibbon, isDialogAvailable } from 'ooxml-ui/pptx';
import type { RibbonCommandRequestEvent } from 'ooxml-ui/pptx';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

import { useResolvedCustomization } from '../../composables/useViewerCustomization';
import type { HelpSectionProps } from './help-section-props';

const props = defineProps<HelpSectionProps>();
const { t } = useI18n();
const customization = useResolvedCustomization();
const groups = computed(() =>
	buildHelpRibbon(t, { options: isDialogAvailable(customization.value, 'options') }),
);
function request(event: Event): void {
	switch ((event as RibbonCommandRequestEvent).detail.id) {
		case 'help.help.options':
			(props.onOpenSettings ?? props.onToggleShortcuts)();
			break;
		case 'help.help.keyboardShortcuts':
			props.onToggleShortcuts();
			break;
		case 'help.help.accessibility':
			props.onRunAccessibilityCheck();
			break;
	}
}
</script>

<template><pptx-ui-ribbon-section :groups.prop="groups" @command-request="request" /></template>
