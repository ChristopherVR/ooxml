<script setup lang="ts">
import type { RibbonTransitionsRequestEvent } from 'ooxml-ui/pptx';
import {
	BUILT_IN_SOUNDS,
	playDeckSound,
	playSlideTransitionPreview,
	readRibbonTransitionDraft,
	ribbonTransitionsIntentUpdates,
	ribbonTransitionsSoundChange,
	ribbonTransitionStockSoundUrl,
} from 'ooxml-ui/pptx';
import type { PptxSlide, PptxSlideTransition } from 'pptx-viewer-core';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

/**
 * TransitionsSection: a thin adapter over the shared `pptx-ui-ribbon-transitions`
 * view. Every control reads the ACTIVE SLIDE through `readRibbonTransitionDraft`
 * and commits the edited fields; Preview replays the transition
 * on the stage without writing, and sound picks are raw transition patches.
 */
interface Props {
	isInspectorPaneOpen: boolean;
	onToggleInspector: () => void;
	canEdit?: boolean;
	/** The slide whose transition the tab reads and writes. */
	activeSlide?: PptxSlide;
	onTransitionChange: (updates: Partial<PptxSlideTransition>) => void;
	onApplyTransitionToAll: () => void;
}

// `withDefaults` is load-bearing: Vue casts an ABSENT boolean prop to `false`,
// so a bare `defineProps` would disable the whole tab whenever a caller omits it.
const props = withDefaults(defineProps<Props>(), { canEdit: true, activeSlide: undefined });
const { t } = useI18n();
const draft = computed(() => readRibbonTransitionDraft(props.activeSlide));
const state = computed(() => ({
	draft: draft.value,
	transition: props.activeSlide?.transition,
	editable: props.canEdit !== false,
	inspectorOpen: props.isInspectorPaneOpen,
	translate: t,
}));

function request(event: RibbonTransitionsRequestEvent): void {
	const intent = event.detail;
	const patch = ribbonTransitionsIntentUpdates(intent, draft.value);
	if (patch) {
		props.onTransitionChange(patch);
		return;
	}
	switch (intent.kind) {
		case 'preview':
			playSlideTransitionPreview(props.activeSlide?.transition, document);
			break;
		case 'applyToAll':
			props.onApplyTransitionToAll();
			break;
		case 'inspector':
			props.onToggleInspector();
			break;
		case 'soundPreview': {
			const url = ribbonTransitionStockSoundUrl(props.activeSlide?.transition);
			if (url) {
				playDeckSound(url, BUILT_IN_SOUNDS);
			}
			break;
		}
		default:
			void ribbonTransitionsSoundChange(intent).then((change) => {
				if (change) {
					props.onTransitionChange(change);
				}
				return undefined;
			});
	}
}
</script>

<template>
	<pptx-ui-ribbon-transitions :state.prop="state" @transitions-request="request" />
</template>
