<script lang="ts">
	/**
	 * RibbonTabBar: the ribbon tab row, the shared `pptx-ui-ribbon-tabs` (the tab row Word, Excel
	 * and Visio draw) fed from `buildRibbonTabRow`, with the right-side quick actions React keeps
	 * on the tab row (`TabRowActions`) in its actions slot: the signature badge, Record, then
	 * Comments and Share drawn by the shared `pptx-ui-ribbon-actions`, then Ribbon Display Options.
	 */
	import {
		buildRibbonTabRow,
		buildTabRowActionsState,
		isActionHidden,
	} from 'ooxml-ui/pptx';
	import {
		RIBBON_DISPLAY_OPTIONS_LABEL_KEY,
	} from 'ooxml-ui/pptx';
	import type {
		RibbonAddInTab,
		RibbonContextualTabId,
		RibbonDisplayController,
		RibbonTabSelectDetail,
		SignatureStatusView,
		ToolbarActionId,
	} from 'ooxml-ui/pptx';
	import { useTranslator } from '../../../i18n/context';
	import type { ChromeUiState } from '../../state/chrome-ui.svelte';
	import { controllerState } from '../../state/controller-state';
	import { useViewerOptions } from '../../state/viewer-options-context';
	import SignatureStatusBadge from '../SignatureStatusBadge.svelte';
	import type { RibbonTabId } from './ribbon-tabs';

	const {
		active,
		onselect,
		onrecord,
		onshare,
		collabActive = false,
		hiddenActions,
		contextualTabs = [],
		addInTabs = [],
		chromeUi,
		commentCount = 0,
		ribbonDisplay,
		signatureView,
		onsignatures,
	}: {
		active: RibbonTabId | RibbonContextualTabId | string;
		onselect: (id: RibbonTabId | RibbonContextualTabId | string) => void;
		/** The host's tabs (`ribbonAddIns`), shown after the fixed tabs. */
		addInTabs?: readonly RibbonAddInTab[];
		/** Selection-driven tabs (Shape Format, ...), appended after the fixed tabs. */
		contextualTabs?: readonly RibbonContextualTabId[];
		onrecord?: () => void;
		onshare?: () => void;
		collabActive?: boolean;
		hiddenActions?: ToolbarActionId[];
		/** Comments pane state and toggle; absent leaves Comments out. */
		chromeUi?: ChromeUiState;
		/** Comments on the current slide, shown as a badge. */
		commentCount?: number;
		/** Ribbon Display Options: the menu button, and a double click on a tab collapses the ribbon. */
		ribbonDisplay?: RibbonDisplayController;
		/** The deck's signature presentation: a badge shows while signature parts are present. */
		signatureView?: SignatureStatusView;
		/** Opens the Digital Signatures dialog from the badge. */
		onsignatures?: () => void;
	} = $props();

	const t = useTranslator();
	const optionsState = useViewerOptions();
	// svelte-ignore state_referenced_locally
	const displayState = ribbonDisplay ? controllerState(ribbonDisplay) : undefined;
	const displayLabel = $derived(t(RIBBON_DISPLAY_OPTIONS_LABEL_KEY));

	const tabRow = $derived(
		buildRibbonTabRow({
			translate: t,
			active,
			hiddenActions,
			addInTabs,
			contextualTabs,
			screenTip: (label) => optionsState.screenTip(label),
		}),
	);
	function onTabSelect(event: Event): void {
		onselect((event as CustomEvent<RibbonTabSelectDetail>).detail.tab);
	}
	const actionsState = $derived(
		buildTabRowActionsState({
			translate: t,
			showComments: Boolean(chromeUi),
			commentsOpen: Boolean(chromeUi?.commentsOpen),
			commentCount,
			showShare: Boolean(onshare) && !isActionHidden('share', hiddenActions),
			isCollaborating: collabActive,
		}),
	);
</script>

<pptx-ui-ribbon-tabs
	data-pptx-chrome="ribbon-tabs"
	state={tabRow}
	ontab-select={onTabSelect}
	oncollapse-toggle={() => ribbonDisplay?.toggleCollapsed()}
>
	<div class="pptx-svelte-ribbon-tabrow-actions" slot="actions">
		{#if signatureView}<SignatureStatusBadge view={signatureView} onclick={onsignatures} />{/if}
		{#if onrecord && !isActionHidden('record', hiddenActions)}
			<button
				type="button"
				class="pptx-svelte-ribbon-record"
				title={t('pptx.titleBar.record')}
				aria-label={t('pptx.titleBar.record')}
				onclick={onrecord}
			>
				<span class="pptx-svelte-ribbon-record-dot" aria-hidden="true"></span>
				<span>{t('pptx.titleBar.record')}</span>
			</button>
		{/if}
		<pptx-ui-ribbon-actions
			data-pptx-chrome="tab-row-actions"
			state={actionsState}
			oncomments-toggle={() => chromeUi?.toggleComments()}
			onshare-request={() => onshare?.()}
		></pptx-ui-ribbon-actions>
		{#if ribbonDisplay && displayState}
			<button
				type="button"
				class="pptx-svelte-ribbon-display-btn"
				class:pptx-svelte-ribbon-display-btn-open={displayState.current.menuOpen}
				class:pptx-svelte-ribbon-display-btn-collapsed={displayState.current.effectiveMode === 'tabs'}
				data-pptx-chrome="ribbon-display-options"
				aria-haspopup="menu"
				aria-expanded={displayState.current.menuOpen}
				aria-label={displayLabel}
				title={displayLabel}
				onclick={() => ribbonDisplay.toggleMenu()}
			>
				<svg viewBox="0 0 16 16" width="12" height="12" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
					<path d="m4 6 4 4 4-4" />
				</svg>
			</button>
		{/if}
	</div>
</pptx-ui-ribbon-tabs>

<style>
	.pptx-svelte-ribbon-tabrow-actions {
		display: flex;
		align-items: center;
		gap: 4px;
		padding-right: 6px;
	}

	.pptx-svelte-ribbon-record {
		display: inline-flex;
		align-items: center;
		gap: 4px;
		padding: 3px 10px;
		border: none;
		border-radius: 4px;
		font: inherit;
		font-size: 11px;
		font-weight: 500;
		white-space: nowrap;
		cursor: pointer;
	}

	.pptx-svelte-ribbon-record {
		background: transparent;
		color: var(--pptx-card-foreground, var(--office-foreground, #1f2937));
	}

	.pptx-svelte-ribbon-record:hover {
		background: var(--pptx-accent, var(--office-selected, #e5e7eb));
	}

	.pptx-svelte-ribbon-record-dot {
		width: 7px;
		height: 7px;
		border-radius: 50%;
		background: #ef4444;
	}

	/* Ribbon Display Options: a small chevron that points the other way when the ribbon is
	   collapsed to its tab row, as PowerPoint's pin does. */
	.pptx-svelte-ribbon-display-btn {
		display: inline-flex;
		align-items: center;
		justify-content: center;
		width: 22px;
		height: 22px;
		padding: 0;
		border: none;
		border-radius: 3px;
		background: transparent;
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
		cursor: pointer;
	}

	.pptx-svelte-ribbon-display-btn:hover,
	.pptx-svelte-ribbon-display-btn-open {
		background: var(--pptx-accent, var(--office-selected, #e5e7eb));
		color: var(--pptx-card-foreground, var(--office-foreground, #1f2937));
	}

	.pptx-svelte-ribbon-display-btn-collapsed svg {
		transform: rotate(180deg);
	}
</style>
