<script setup lang="ts">
/**
 * ContextualTabSection: the body of a contextual ribbon tab (Shape Format,
 * Picture Format, Table Design, Chart Design, SmartArt Design). The groups
 * and the galleries inside them, with their modes, come from the shared
 * `CONTEXTUAL_TAB_GROUPS`; this only maps them onto the ribbon's group markup.
 * A command group (Picture Format > Size, Chart Format > Current Selection)
 * draws its shared strip as well.
 */
import { CONTEXTUAL_TAB_GROUPS } from 'ooxml-ui/pptx';
import type { RibbonContextualTabId, ToolbarActionId } from 'ooxml-ui/pptx';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

import ChartSelectionStrip from './ChartSelectionStrip.vue';
import PictureSizeStrip from './PictureSizeStrip.vue';
import RibbonGallery from './RibbonGallery.vue';

interface Props {
	tab: RibbonContextualTabId;
	/** Host + customisation hidden actions (gates the Size group's Crop). */
	hiddenActions?: readonly ToolbarActionId[];
}

const props = defineProps<Props>();
const { t } = useI18n();
const groups = computed(() => CONTEXTUAL_TAB_GROUPS[props.tab]);

function caption(labelKey: string, fallback: string): string {
	const translated = t(labelKey);
	return translated && translated !== labelKey ? translated : fallback;
}
</script>

<template>
	<pptx-ui-ribbon-group
		v-for="group in groups"
		:key="group.group"
		:label="caption(group.labelKey, group.label)"
		:data-ribbon-group="group.group"
	>
		<RibbonGallery
			v-for="placement in group.galleries"
			:key="placement.control"
			:gallery="placement.gallery"
			:control="placement.control"
			:mode="placement.mode"
		/>
		<PictureSizeStrip v-if="group.strip === 'picture-size'" :hidden-actions="props.hiddenActions" />
		<ChartSelectionStrip v-else-if="group.strip === 'chart-selection'" />
	</pptx-ui-ribbon-group>
</template>
