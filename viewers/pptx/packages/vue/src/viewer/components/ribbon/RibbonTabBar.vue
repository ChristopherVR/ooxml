<script setup lang="ts">
/**
 * RibbonTabBar: the ribbon's tab row, the shared `pptx-ui-ribbon-tabs` (the tab row Word, Excel
 * and Visio draw) fed from `buildRibbonTabRow`, with the Record + Share cluster and the Ribbon
 * Display Options (or collapse) button in its actions slot. Extracted from `RibbonToolbar.vue`;
 * the props are the ones `RibbonToolbar` already reads off `RibbonProps`.
 */
import { ChevronDown, ChevronUp } from '@lucide/vue';
import { buildRibbonTabRow, RIBBON_DISPLAY_OPTIONS_LABEL_KEY } from 'ooxml-ui/pptx';
import type {
	RibbonAddInTab,
	RibbonContextualTabId,
	RibbonDisplayController,
	RibbonDisplayState,
	RibbonTabSelectDetail,
	ToolbarActionId,
	ToolbarTabDefinition,
} from 'ooxml-ui/pptx';
import { computed, inject } from 'vue';
import { useI18n } from 'vue-i18n';

import { cn } from '../../../utils';
import { ScreenTipKey } from '../../composables/useViewerOptionsStore';
import type { ToolbarSection, ViewerMode } from './ribbon-types';
import TabRowActions from './TabRowActions.vue';

interface Props {
	toolbarSection: ToolbarSection;
	visibleTabs: ToolbarTabDefinition[];
	/** Contextual tabs the selection brings up (shared `visibleContextualTabs`). */
	contextualTabs?: readonly RibbonContextualTabId[];
	/** The host's tabs (`ribbonAddIns`), shown after the fixed tabs. */
	addInTabs?: readonly RibbonAddInTab[];
	/** The host tab being shown; no fixed or contextual tab is selected while one is. */
	activeAddIn?: string | null;
	onSelectAddIn?: (id: string) => void;
	onSetToolbarSection: (section: ToolbarSection) => void;
	canEdit: boolean;
	onEnterRehearsalMode?: () => void;
	onSetMode: (mode: ViewerMode) => void;
	onOpenShareDialog?: () => void;
	isCollaborating?: boolean;
	collaboratorCount?: number;
	onToggleComments?: (ensureOpen?: boolean) => void;
	isCommentsPanelOpen?: boolean;
	slideCommentCount?: number;
	hiddenActions?: ToolbarActionId[];
	isCompactToolbarOpen: boolean;
	onToggleCompactToolbar: () => void;
	/** Ribbon Display Options state (absent: the menu button is not drawn). */
	ribbonDisplay?: RibbonDisplayState;
	ribbonDisplayController?: RibbonDisplayController;
}

const props = defineProps<Props>();
const { t } = useI18n();
const screenTip = inject(ScreenTipKey, (label: string) => label);

/** The tab row as the shared element draws it; `visibleTabs` is already filtered. */
const tabRow = computed(() =>
	buildRibbonTabRow({
		translate: (key) => t(key),
		tabs: props.visibleTabs,
		active: props.activeAddIn ?? props.toolbarSection,
		addInTabs: props.addInTabs,
		contextualTabs: props.contextualTabs,
		screenTip,
	}),
);

/** A tab was chosen: the binding's own switch runs first, then the display controller peeks. */
function onTabSelect(event: Event): void {
	const { tab, kind } = (event as CustomEvent<RibbonTabSelectDetail>).detail;
	if (kind === 'addIn') {
		props.onSelectAddIn?.(tab);
	} else {
		props.onSetToolbarSection(tab as ToolbarSection);
	}
	props.ribbonDisplayController?.tabChosen(tab);
}
</script>

<template>
	<pptx-ui-ribbon-tabs
		data-pptx-chrome="ribbon-tabs"
		:state="tabRow"
		@tab-select="onTabSelect"
		@collapse-toggle="props.ribbonDisplayController?.toggleCollapsed()"
	>
		<div slot="actions" class="flex items-center">
			<TabRowActions
				:on-enter-rehearsal-mode="
					props.canEdit
						? (props.onEnterRehearsalMode ?? (() => props.onSetMode('present')))
						: undefined
				"
				:on-open-share-dialog="props.onOpenShareDialog"
				:is-collaborating="props.isCollaborating"
				:collaborator-count="props.collaboratorCount"
				:on-toggle-comments="props.onToggleComments"
				:is-comments-panel-open="props.isCommentsPanelOpen"
				:slide-comment-count="props.slideCommentCount"
				:hidden-actions="props.hiddenActions"
			/>
			<button
				v-if="props.ribbonDisplayController"
				type="button"
				data-pptx-chrome="ribbon-display-options"
				aria-haspopup="menu"
				:aria-expanded="props.ribbonDisplay?.menuOpen ?? false"
				:aria-label="t(RIBBON_DISPLAY_OPTIONS_LABEL_KEY)"
				:title="t(RIBBON_DISPLAY_OPTIONS_LABEL_KEY)"
				:class="
					cn(
						'mr-1 inline-flex items-center justify-center rounded px-1.5 py-1 transition-colors',
						props.ribbonDisplay?.menuOpen
							? 'bg-primary text-white'
							: 'text-muted-foreground hover:text-foreground hover:bg-accent/30',
					)
				"
				@click="props.ribbonDisplayController?.toggleMenu()"
			>
				<ChevronDown
					:class="[
						'h-3.5 w-3.5 transition-transform',
						props.ribbonDisplay?.effectiveMode === 'tabs' ? 'rotate-180' : '',
					]"
					aria-hidden="true"
				/>
			</button>
			<!-- The plain collapse toggle of a tab row used without Ribbon Display Options; the
		     menu above replaces it in the viewer (one arrow, as in PowerPoint). -->
			<button
				v-else
				type="button"
				class="mr-1 inline-flex items-center justify-center rounded px-2 py-1 text-muted-foreground transition-colors hover:text-foreground"
				:aria-pressed="!props.isCompactToolbarOpen"
				:aria-label="
					props.isCompactToolbarOpen
						? t('pptx.ribbon.collapseRibbon')
						: t('pptx.ribbon.expandRibbon')
				"
				:title="
					props.isCompactToolbarOpen
						? t('pptx.ribbon.collapseRibbon')
						: t('pptx.ribbon.expandRibbon')
				"
				@click="props.onToggleCompactToolbar"
			>
				<component
					:is="props.isCompactToolbarOpen ? ChevronUp : ChevronDown"
					class="h-3.5 w-3.5"
					aria-hidden="true"
				/>
			</button>
		</div>
	</pptx-ui-ribbon-tabs>
</template>
