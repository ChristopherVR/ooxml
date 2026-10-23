<script setup lang="ts">
/**
 * InsertSection: thin adapter for the shared `pptx-ui-ribbon-insert`. The shared
 * element owns the groups, icons, labels, pickers and pressed/disabled state; this
 * component supplies viewer state and routes typed intents to the native handlers.
 * Document mutation, the file/SmartArt/equation/hyperlink dialogs and the Date/Time
 * picker stay native to the Vue binding.
 */
import {
	DEFAULT_INSERT_CHART_KIND,
	FREEFORM_TOOL_IDS,
	isDialogAvailable,
	isDrawingToolVisible,
	resolveInsertSlideLayout,
} from 'ooxml-ui/pptx';
import type {
	FreeformToolKind,
	InsertChartKind,
	InsertChartOptions,
	InsertSlideLayout,
	RibbonInsertRequestEvent,
	SlideTemplateId,
} from 'ooxml-ui/pptx';
import { computed, ref } from 'vue';
import { useI18n } from 'vue-i18n';

import { useOutlineAuthoring } from '../../composables/useOutlineAuthoring';
import { useResolvedCustomization } from '../../composables/useViewerCustomization';
import SlideTemplateGalleryDialog from '../SlideTemplateGalleryDialog.vue';
import DateTimeFieldDialog from './DateTimeFieldDialog.vue';
import InsertChartDialog from './InsertChartDialog.vue';
import type { SupportedShapeType } from './ribbon-types';

interface Props {
	canEdit: boolean;
	/** Whether an element is selected; gates the Link button. */
	hasSelection: boolean;
	onOpenHyperlinkDialog: () => void;
	newShapeType: SupportedShapeType;
	onSetNewShapeType: (type: SupportedShapeType) => void;
	onAddTextBox: () => void;
	onAddShape: () => void;
	onAddTable: (rows?: number, columns?: number) => void;
	onAddChart?: (chartKind: InsertChartKind, options?: InsertChartOptions) => void;
	onAddSmartArt: () => void;
	onAddEquation: () => void;
	onAddActionButton: (shapeType: string) => void;
	onInsertField?: (fieldType: string, value?: string) => void;
	onOpenHeaderFooter?: () => void;
	onOpenImagePicker: () => void;
	onOpenMediaPicker: () => void;
	/** The layouts New Slide offers, and the active slide's layout. */
	layoutOptions?: readonly InsertSlideLayout[];
	currentLayoutPath?: string;
	onInsertSlideFromLayout?: (path: string, name?: string) => void;
	/** Slide Templates... ends New Slide's menu; without a handler the entry is left out. */
	onInsertSlideFromTemplate?: (templateId: SlideTemplateId) => void;
	/** Deck scheme map so template previews show the deck's theme colours. */
	templateScheme?: Record<string, string>;
	/** Comment opens the comments pane, as Review > New Comment does. */
	onToggleComments?: () => void;
	isCommentsPanelOpen?: boolean;
	/** Draw the phone Insert sheet (`pptx-ui-mobile-insert`) instead of the ribbon group row. */
	phone?: boolean;
	/** Phone only: called after a request is handled, so the sheet can close. */
	onRequestHandled?: () => void;
	/**
	 * Phone only: the Date & Time and Slide Templates dialogs are owned by a host that outlives
	 * the sheet (closing the sheet unmounts this section and would take them with it).
	 */
	dialogHost?: InsertDialogHost;
}

export interface InsertDialogHost {
	openDateTime: () => void;
	openSlideTemplates: () => void;
	/** Opens Insert > Chart > Insert Chart... (the shared dialog) outside the sheet. */
	openInsertChart?: () => void;
}

const props = defineProps<Props>();
const { t } = useI18n();
const outline = useOutlineAuthoring();
const customization = useResolvedCustomization();
const chartKind = ref<InsertChartKind>(DEFAULT_INSERT_CHART_KIND);
const datePickerOpen = ref(false);
const templatesOpen = ref(false);
const chartDialogOpen = ref(false);
const templatesAvailable = computed(
	() =>
		Boolean(props.onInsertSlideFromTemplate) &&
		isDialogAvailable(customization.value, 'slideTemplates'),
);

const state = computed(() => ({
	editable: props.canEdit,
	hasSelection: props.hasSelection,
	shapeType: props.newShapeType,
	chartKind: chartKind.value,
	activeFreeformTool: outline?.activeFreeformTool.value ?? null,
	// Renders nothing outside a viewer (no outline-authoring store provided).
	freeformTools: outline
		? FREEFORM_TOOL_IDS.filter((tool) => isDrawingToolVisible(customization.value, tool))
		: [],
	chartAvailable: Boolean(props.onAddChart),
	chartDialogAvailable: Boolean(props.onAddChart),
	fieldAvailable: Boolean(props.onInsertField),
	headerFooterAvailable: Boolean(props.onOpenHeaderFooter),
	newSlideAvailable: Boolean(props.onInsertSlideFromLayout && props.layoutOptions?.length),
	commentAvailable: Boolean(props.onToggleComments),
	slideTemplatesAvailable: templatesAvailable.value,
	layouts: props.layoutOptions ?? [],
	...(props.currentLayoutPath ? { currentLayoutPath: props.currentLayoutPath } : {}),
	translate: t,
}));

function request(event: RibbonInsertRequestEvent): void {
	const intent = event.detail;
	switch (intent.kind) {
		case 'command': {
			const commands = {
				textBox: props.onAddTextBox,
				table: props.onAddTable,
				image: props.onOpenImagePicker,
				media: props.onOpenMediaPicker,
				smartArt: props.onAddSmartArt,
				equation: props.onAddEquation,
				link: props.onOpenHyperlinkDialog,
				headerFooter: props.onOpenHeaderFooter,
				// Opening an already open pane would toggle it shut.
				comment: props.isCommentsPanelOpen ? undefined : props.onToggleComments,
			};
			commands[intent.value]?.();
			break;
		}
		case 'table':
			props.onAddTable(intent.rows, intent.columns);
			break;
		case 'newSlide': {
			const layout = resolveInsertSlideLayout(
				props.layoutOptions ?? [],
				intent.value,
				props.currentLayoutPath,
			);
			if (layout) {
				props.onInsertSlideFromLayout?.(layout.path, layout.name);
			}
			break;
		}
		case 'slideTemplates':
			if (props.dialogHost) {
				props.dialogHost.openSlideTemplates();
			} else {
				templatesOpen.value = true;
			}
			break;
		case 'shapeType':
			props.onSetNewShapeType(intent.value as SupportedShapeType);
			break;
		case 'shape':
			props.onAddShape();
			break;
		case 'chartType':
			chartKind.value = intent.value as InsertChartKind;
			break;
		case 'chart':
			props.onAddChart?.(intent.value as InsertChartKind);
			break;
		case 'chartDialog':
			if (props.dialogHost?.openInsertChart) {
				props.dialogHost.openInsertChart();
			} else {
				chartDialogOpen.value = true;
			}
			break;
		case 'freeform':
			outline?.armFreeformTool(intent.value as FreeformToolKind | null);
			break;
		case 'actionButton':
			props.onAddActionButton(intent.value);
			break;
		case 'field':
			if (intent.value === 'datetime') {
				if (props.dialogHost) {
					props.dialogHost.openDateTime();
				} else {
					datePickerOpen.value = true;
				}
			} else {
				props.onInsertField?.(intent.value);
			}
	}
	props.onRequestHandled?.();
}

function insertDate(formatted: string): void {
	props.onInsertField?.('datetime', formatted);
	datePickerOpen.value = false;
}
</script>

<template>
	<pptx-ui-mobile-insert v-if="props.phone" :state.prop="state" @insert-request="request" />
	<pptx-ui-ribbon-insert v-else :state.prop="state" @insert-request="request" />
	<SlideTemplateGalleryDialog
		v-if="templatesAvailable && !props.dialogHost"
		:open="templatesOpen"
		:scheme="props.templateScheme"
		@insert="props.onInsertSlideFromTemplate?.($event)"
		@close="templatesOpen = false"
	/>
	<InsertChartDialog
		v-if="props.onAddChart && !props.dialogHost?.openInsertChart"
		:open="chartDialogOpen"
		:on-add-chart="props.onAddChart"
		@close="chartDialogOpen = false"
	/>
	<DateTimeFieldDialog
		v-if="datePickerOpen && !props.dialogHost && props.onInsertField"
		@close="datePickerOpen = false"
		@insert="insertDate"
	/>
</template>
