<script setup lang="ts">
/**
 * MobileInsertSheet: the bottom bar's Insert sheet. The shared `pptx-ui-mobile-insert` tile
 * grid, routed through the same `InsertSection` handling as the ribbon, so a pick runs the
 * identical command (React: `MobileInsertSheet.tsx`).
 */
import { isDialogAvailable } from 'ooxml-ui/pptx';
import { computed, ref } from 'vue';
import { useI18n } from 'vue-i18n';

import { useResolvedCustomization } from '../composables/useViewerCustomization';
import MobileSheet from './MobileSheet.vue';
import DateTimeFieldDialog from './ribbon/DateTimeFieldDialog.vue';
import InsertSection from './ribbon/InsertSection.vue';
import type { RibbonProps } from './ribbon/ribbon-types';
import SlideTemplateGalleryDialog from './SlideTemplateGalleryDialog.vue';

interface Props extends RibbonProps {
	open: boolean;
}

const props = defineProps<Props>();
const emit = defineEmits<{ close: [] }>();
const { t } = useI18n();
const customization = useResolvedCustomization();
// These dialogs are opened from the sheet but must outlive it: the sheet unmounts its body the
// moment a pick closes it, which would take a dialog owned by the section along.
const dateOpen = ref(false);
const templatesOpen = ref(false);
const templatesAvailable = computed(
	() =>
		Boolean(props.onInsertSlideFromTemplate) &&
		isDialogAvailable(customization.value, 'slideTemplates'),
);
const dialogHost = {
	openDateTime: () => {
		dateOpen.value = true;
	},
	openSlideTemplates: () => {
		templatesOpen.value = true;
	},
};
function insertDate(formatted: string): void {
	props.onInsertField?.('datetime', formatted);
	dateOpen.value = false;
}
</script>

<template>
	<MobileSheet :open="props.open" :title="t('pptx.mobileBar.insert')" @close="emit('close')">
		<div class="p-3">
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
				:on-insert-slide-from-template="props.onInsertSlideFromTemplate"
				:template-scheme="props.templateScheme"
				:on-toggle-comments="props.onToggleComments"
				:is-comments-panel-open="props.isCommentsPanelOpen"
				:dialog-host="dialogHost"
			/>
		</div>
	</MobileSheet>
	<SlideTemplateGalleryDialog
		v-if="templatesAvailable"
		:open="templatesOpen"
		:scheme="props.templateScheme"
		@insert="props.onInsertSlideFromTemplate?.($event)"
		@close="templatesOpen = false"
	/>
	<DateTimeFieldDialog
		v-if="dateOpen && props.onInsertField"
		@close="dateOpen = false"
		@insert="insertDate"
	/>
</template>
