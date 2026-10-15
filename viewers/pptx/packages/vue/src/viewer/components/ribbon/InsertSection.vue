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
	isDrawingToolVisible,
	resolveInsertSlideLayout,
} from 'ooxml-ui/pptx';
import type {
	FreeformToolKind,
	InsertChartKind,
	InsertSlideLayout,
	RibbonInsertRequestEvent,
} from 'ooxml-ui/pptx';
import { computed, ref } from 'vue';
import { useI18n } from 'vue-i18n';

import { useOutlineAuthoring } from '../../composables/useOutlineAuthoring';
import { useResolvedCustomization } from '../../composables/useViewerCustomization';
import DateTimeFieldDialog from './DateTimeFieldDialog.vue';
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
	onAddChart?: (chartKind: InsertChartKind) => void;
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
	/** Comment opens the comments pane, as Review > New Comment does. */
	onToggleComments?: () => void;
	isCommentsPanelOpen?: boolean;
	/** Draw the phone Insert sheet (`pptx-ui-mobile-insert`) instead of the ribbon group row. */
	phone?: boolean;
	/** Phone only: called after a request is handled, so the sheet can close. */
	onRequestHandled?: () => void;
}

const props = defineProps<Props>();
const { t } = useI18n();
const outline = useOutlineAuthoring();
const customization = useResolvedCustomization();
const chartKind = ref<InsertChartKind>(DEFAULT_INSERT_CHART_KIND);
const datePickerOpen = ref(false);

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
	fieldAvailable: Boolean(props.onInsertField),
	headerFooterAvailable: Boolean(props.onOpenHeaderFooter),
	newSlideAvailable: Boolean(props.onInsertSlideFromLayout && props.layoutOptions?.length),
	commentAvailable: Boolean(props.onToggleComments),
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
		case 'freeform':
			outline?.armFreeformTool(intent.value as FreeformToolKind | null);
			break;
		case 'actionButton':
			props.onAddActionButton(intent.value);
			break;
		case 'field':
			if (intent.value === 'datetime') {
				datePickerOpen.value = true;
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
	<DateTimeFieldDialog
		v-if="datePickerOpen && props.onInsertField"
		@close="datePickerOpen = false"
		@insert="insertDate"
	/>
</template>
