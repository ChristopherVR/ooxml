<script setup lang="ts">
/**
 * MobileInsertSheet: the bottom bar's Insert sheet. The shared `pptx-ui-mobile-insert` tile
 * grid, routed through the same `InsertSection` handling as the ribbon, so a pick runs the
 * identical command (React: `MobileInsertSheet.tsx`).
 */
import { useI18n } from 'vue-i18n';

import MobileSheet from './MobileSheet.vue';
import InsertSection from './ribbon/InsertSection.vue';
import type { RibbonProps } from './ribbon/ribbon-types';

interface Props extends RibbonProps {
	open: boolean;
}

const props = defineProps<Props>();
const emit = defineEmits<{ close: [] }>();
const { t } = useI18n();
</script>

<template>
	<MobileSheet :open="props.open" :title="t('pptx.mobileBar.insert')" @close="emit('close')">
		<div class="p-3 pb-20">
			<InsertSection
				phone
				:on-request-handled="() => emit('close')"
				:can-edit="props.canEdit"
				:new-shape-type="props.newShapeType"
				:on-set-new-shape-type="props.onSetNewShapeType"
				:on-add-text-box="props.onAddTextBox"
				:on-add-shape="props.onAddShape"
				:on-add-table="props.onAddTable"
				:on-add-chart="props.onAddChart"
				:on-add-smart-art="props.onAddSmartArt"
				:on-add-equation="props.onAddEquation"
				:on-add-action-button="props.onAddActionButton"
				:on-insert-field="props.onInsertField"
				:on-open-header-footer="props.onOpenHeaderFooter"
				:on-open-image-picker="props.onOpenImagePicker"
				:on-open-media-picker="props.onOpenMediaPicker"
				:has-selection="Boolean(props.selectedElement)"
				:on-open-hyperlink-dialog="props.onOpenHyperlinkDialog"
				:layout-options="props.layoutOptions"
				:current-layout-path="props.currentLayoutPath"
				:on-insert-slide-from-layout="props.onInsertSlideFromLayout"
				:on-toggle-comments="props.onToggleComments"
				:is-comments-panel-open="props.isCommentsPanelOpen"
			/>
		</div>
	</MobileSheet>
</template>
