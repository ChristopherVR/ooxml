<script setup lang="ts">
/**
 * StatusBar - thin adapter around the shared `pptx-ui-status-bar` element.
 *
 * The slide counter, save indicator, notes toggle, view-mode buttons and zoom
 * cluster live in the shared view. This component maps viewer state onto the
 * element's controlled state and re-emits its typed intents.
 */
import { dispatchStatusBarIntent, statusBarState, statusBarViewMode } from 'ooxml-ui/pptx';
import type { ToolbarActionId } from 'ooxml-ui/pptx';
import { computed, useSlots } from 'vue';
import { useI18n } from 'vue-i18n';

import type { AutosaveStatus } from '../composables/useAutosave';
import { useRegion } from '../composables/useViewerRegions';
import ViewerRegion from './ViewerRegion.vue';

const { t } = useI18n();

const props = defineProps<{
	slideCount: number;
	activeSlideIndex: number;
	isDirty: boolean;
	autosaveStatus?: AutosaveStatus;
	/** Epoch ms of the last successful autosave, for the "Saved <time>" label. */
	lastSavedAt?: number | null;
	/** Current zoom scale (1 = 100%). */
	scale?: number;
	/** Whether the notes panel is expanded. */
	isNotesExpanded?: boolean;
	/** Current viewer mode. */
	mode?: string;
	/** Whether the Notes toggle is shown (host has a notes panel). */
	showNotes?: boolean;
	/** Toolbar buttons the host has asked to hide (zoom, notes, fullscreen). */
	hiddenActions?: ToolbarActionId[];
}>();

// A host `regions.statusBar` replaces the whole bar in its slot.
const region = useRegion('statusBar');
const regionContext = computed(() => region.context(() => undefined));

const slots = useSlots();

const emit = defineEmits<{
	'zoom-in': [];
	'zoom-out': [];
	'zoom-to-fit': [];
	'toggle-notes': [];
	'set-mode': [mode: 'edit' | 'present'];
	'toggle-slide-sorter': [];
}>();

const state = computed(() =>
	statusBarState({
		slideCount: props.slideCount,
		activeSlideIndex: props.activeSlideIndex,
		dirty: props.isDirty,
		autosave: props.autosaveStatus
			? { state: props.autosaveStatus, timestamp: props.lastSavedAt ?? undefined }
			: undefined,
		zoomPercent: props.scale === undefined ? undefined : props.scale * 100,
		notesExpanded: props.isNotesExpanded,
		viewMode: statusBarViewMode(props.mode),
		handles: { notes: props.showNotes === true },
		hiddenActions: props.hiddenActions,
		translate: t,
	}),
);

function request(event: Event): void {
	dispatchStatusBarIntent(event, {
		notes: () => emit('toggle-notes'),
		normal: () => emit('set-mode', 'edit'),
		sorter: () => emit('toggle-slide-sorter'),
		slideShow: () => emit('set-mode', 'present'),
		zoomOut: () => emit('zoom-out'),
		zoomFit: () => emit('zoom-to-fit'),
		zoomIn: () => emit('zoom-in'),
	});
}
</script>

<template>
	<div
		v-if="region.renderer.value && regionContext"
		data-pptx-chrome="status-bar"
		style="flex: none"
	>
		<ViewerRegion :renderer="region.renderer.value" :context="regionContext" />
	</div>
	<pptx-ui-status-bar v-else :state.prop="state" @status-request="request">
		<div v-if="slots.collaboration" slot="collaboration" style="display: contents">
			<slot name="collaboration" />
		</div>
	</pptx-ui-status-bar>
</template>
