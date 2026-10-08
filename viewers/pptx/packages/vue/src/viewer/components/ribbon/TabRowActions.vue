<script setup lang="ts">
/**
 * TabRowActions: Vue port of React's `toolbar/TabRowActions.tsx`.
 *
 * Right-side actions on the ribbon tab row: Record (starts rehearsal mode,
 * records slide timings), then Comments and Share, drawn by the shared
 * `pptx-ui-ribbon-actions` (the element Word and Excel put at the same place).
 * Share reads pressed while a collaboration session is connected.
 *
 * React reads the collaboration state from a `useCollaboration()` context; in
 * Vue collaboration is host-instantiated, so the connected state is threaded in
 * as `isCollaborating` / `collaboratorCount` props (surfaced through
 * `RibbonProps`).
 */
import {
	buildTabRowActionsState,
	isFeatureEnabled,
	TAB_ROW_ACTION_CLASSES as TRA,
} from 'ooxml-ui/pptx';
import type { ToolbarActionId } from 'ooxml-ui/pptx';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

import { useToolbarVisibility } from '../../composables/useToolbarVisibility';
import { useResolvedCustomization } from '../../composables/useViewerCustomization';

interface Props {
	onEnterRehearsalMode?: () => void;
	onOpenShareDialog?: () => void;
	isCollaborating?: boolean;
	collaboratorCount?: number;
	/** Comments toggle; absent (or the host turning comments off) leaves Comments out. */
	onToggleComments?: () => void;
	isCommentsPanelOpen?: boolean;
	slideCommentCount?: number;
	/** Toolbar buttons the host has asked to hide (gates Record + Share below). */
	hiddenActions?: ToolbarActionId[];
}

const props = defineProps<Props>();
const { t } = useI18n();
const { isHidden } = useToolbarVisibility(() => props.hiddenActions);
const customization = useResolvedCustomization();
const state = computed(() =>
	buildTabRowActionsState({
		translate: (key, params) => (params ? t(key, params) : t(key)),
		showComments:
			Boolean(props.onToggleComments) && isFeatureEnabled(customization.value, 'comments'),
		commentsOpen: Boolean(props.isCommentsPanelOpen),
		commentCount: props.slideCommentCount ?? 0,
		showShare: !isHidden('share'),
		isCollaborating: Boolean(props.isCollaborating),
		collaboratorCount: props.collaboratorCount ?? 0,
	}),
);
</script>

<template>
	<div class="flex items-center gap-1 pr-1">
		<button
			v-if="props.onEnterRehearsalMode && !isHidden('record')"
			type="button"
			:class="TRA.record"
			:title="t('pptx.titleBar.record')"
			:aria-label="t('pptx.titleBar.record')"
			@click="props.onEnterRehearsalMode()"
		>
			<span :class="TRA.recordDot" aria-hidden="true" />
			<span>{{ t('pptx.titleBar.record') }}</span>
		</button>
		<pptx-ui-ribbon-actions
			data-pptx-chrome="tab-row-actions"
			:state.prop="state"
			@comments-toggle="props.onToggleComments?.()"
			@share-request="props.onOpenShareDialog?.()"
		/>
	</div>
</template>
