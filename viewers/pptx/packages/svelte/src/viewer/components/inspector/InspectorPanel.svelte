<script lang="ts">
	/**
	 * InspectorPanel: right-hand inspector pane, structured like React's
	 * `InspectorPane`: an [Elements | Properties | Comments] tab strip with a
	 * close button, defaulting to Properties.
	 *
	 * - Elements: the active slide's layer-order list ({@link ElementsListSection}).
	 * - Properties: a selection-aware Format pane ({@link FormatSections}): with
	 *   a selection, the shared collapsible sections for its kind
	 *   ({@link ElementFormatSections}); with no selection, one Slide section
	 *   holding the deck-level sections in React's order
	 *   ({@link PresentationPropertiesPanel}: PRESENTATION, THEME, THEME
	 *   OVERRIDE, SLIDE SIZE, NOTES & HANDOUT, DOCUMENT).
	 * - Comments: the slide's comment thread ({@link ReviewCommentsPanel}).
	 *
	 * Every control routes edits through `EditorState.applyElementPatch` /
	 * `patchSelected`, so every change is undo/redo-integrated. Open state and
	 * the active tab live in {@link ChromeUiState} when the host passes `ui`
	 * (so the toolbar's comments/inspector toggles stay in sync); standalone
	 * mounts fall back to local state.
	 */
	import { Layers, Settings2 } from '@lucide/svelte';
	import type { PptxElement, PptxHandler, PptxTheme } from 'pptx-viewer-core';
	import { createFormatPaneController, getDensePanelTouchTargetPx } from 'ooxml-ui/pptx';
	import type { CanvasSize } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';
	import { useViewerCustomization } from '../../state/viewer-customization.svelte';
	import ViewerRegion from '../ViewerRegion.svelte';
	import type { EditorState } from '../../editor/editor-state.svelte';
	import type { ChromeUiState, InspectorTabId } from '../../state/chrome-ui.svelte';
	import { useInspectorDeck } from '../../state/inspector-deck';
	import { useWindowViewport } from '../../state/window-viewport.svelte';
	import { attachCommentsTaskPane } from 'ooxml-ui/pptx';
	import ReviewCommentsPanel from '../ribbon/review/ReviewCommentsPanel.svelte';
	import AnimationPanel from './AnimationPanel.svelte';
	import ElementFormatSections from './ElementFormatSections.svelte';
	import ElementsListSection from './ElementsListSection.svelte';
	import FormatSections from './FormatSections.svelte';
	import PresentationPropertiesPanel from './PresentationPropertiesPanel.svelte';

	const { editor, handler, presentationTheme, onthemechange, mediaDataUrls = new Map(), ui, canvasSize }: { editor: EditorState; handler?: PptxHandler | null; presentationTheme?: PptxTheme; onthemechange?: (theme: PptxTheme) => void; mediaDataUrls?: Map<string, string>; ui?: ChromeUiState; canvasSize?: CanvasSize } = $props();
	const t = useTranslator();
	// Host-supplied inspector content (`regions.inspector`): replaces the pane, keeps the column.
	const customization = useViewerCustomization();
	const regionRenderer = $derived(customization.regionRenderer('inspector'));
	const viewport = useWindowViewport();
	// The tab-strip buttons are sized for a mouse (3px/8px padding); below the
	// dense-panel breakpoint they must clear the 44px WCAG touch target.
	const tabTouchPx = $derived(getDensePanelTouchTargetPx(viewport.width));
	const tabTouchStyle = $derived(`min-width: ${tabTouchPx}px; min-height: ${tabTouchPx}px;`);
	// Deck-level state/mutations for the no-selection Properties tab, provided
	// by `PowerPointViewer` (undefined in standalone mounts, e.g. tests).
	const deck = useInspectorDeck();

	// Standalone fallbacks when no ChromeUiState is provided (tests, hosts).
	let localTab = $state<InspectorTabId>('properties');
	const activeTab = $derived(ui ? ui.inspectorTab : localTab);
	function setTab(tab: InspectorTabId): void {
		if (ui) {
			ui.setInspectorTab(tab);
		} else {
			localTab = tab;
		}
	}

	// Same dictionary key the vanilla inspector uses for its Elements tab.
	const tabs = $derived<Array<{ id: InspectorTabId; label: string }>>([
		{ id: 'elements', label: t('pptx.documentProperties.statistics.elements') },
		{ id: 'properties', label: t('pptx.inspector.properties') },
	]);

	const el = $derived(editor.selectedElement);
	const activeSlide = $derived(editor.slides[editor.currentSlideIndex]);
	// The viewer's Format pane controller; a standalone mount keeps its own.
	const localPane = createFormatPaneController({ showPane: () => setTab('properties') });
	const pane = $derived(ui?.formatPane ?? localPane);
	const NO_SELECTION: readonly PptxElement[] = [];
	function commentsTaskPane(node: HTMLElement) {
		return { destroy: attachCommentsTaskPane(node, {
			label: t('pptx.toolbar.comments'),
			closeLabel: t('pptx.common.close'),
			close: () => ui?.toggleComments(),
		}) };
	}
</script>

<aside
	class="pptx-svelte-inspector"
	data-pptx-inspector data-pptx-chrome="inspector"
	aria-label={t('pptx.inspector.properties')}
>
	{#if regionRenderer}
		<ViewerRegion region="inspector" renderer={regionRenderer} close={() => (ui ? (ui.inspectorOpen = false) : undefined)} />
	{:else if activeTab === 'comments'}
		<div class="pptx-svelte-inspector-body" role="complementary" use:commentsTaskPane>
			<ReviewCommentsPanel {editor} />
		</div>
	{:else}
	<div class="pptx-svelte-inspector-header" data-pptx-chrome="inspector-header">
		<div class="pptx-svelte-inspector-tabs" data-pptx-chrome="inspector-tabs" role="tablist">
			{#each tabs as tab (tab.id)}
				<button
					type="button"
					role="tab"
					style={tabTouchStyle}
					aria-selected={activeTab === tab.id}
					class:pptx-svelte-inspector-tab-active={activeTab === tab.id}
					onclick={() => setTab(tab.id)}
				>
					{#if tab.id === 'elements'}<Layers size={14} aria-hidden="true" />{:else if tab.id === 'properties'}<Settings2 size={14} aria-hidden="true" />{/if}
					{tab.label}
				</button>
			{/each}
		</div>
	</div>

	<div class="pptx-svelte-inspector-body" data-pptx-chrome="inspector-body">
			{#if activeTab === 'elements'}
				<ElementsListSection {editor} />

			{:else if el}
				<ElementFormatSections {editor} {el} {pane} {deck} {mediaDataUrls} />
			{:else}
				<FormatSections elements={NO_SELECTION} {pane}>
					{#snippet section()}
						<PresentationPropertiesPanel {editor} {deck} {canvasSize} {handler} {presentationTheme} {onthemechange} />
						{#if !activeSlide}
							<!-- React parity: only when there is genuinely no active slide. -->
							<p class="pptx-svelte-inspector-empty">{t('pptx.inspector.noSlideSelected')}</p>
						{/if}
					{/snippet}
				</FormatSections>
			{/if}
	</div>

	<!-- Bottom dock: per-element animations, any tab (React InspectorPane parity). -->
	<AnimationPanel {editor} />
	{/if}
</aside>

<style>
	.pptx-svelte-inspector {
		display: flex;
		flex-direction: column;
		width: 288px;
		flex: none;
		border-left: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		background: var(--pptx-card, var(--office-surface, #f3f4f6));
		color: var(--pptx-card-foreground, var(--office-foreground, #1f2937));
		font-family: system-ui, sans-serif;
		font-size: 12px;
		overflow: hidden;
		min-height: 0;
	}

	.pptx-svelte-inspector-header {
		display: flex;
		align-items: center;
		gap: 8px;
		padding: 8px 10px;
		border-bottom: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
	}

	.pptx-svelte-inspector-tabs {
		display: inline-flex;
		align-items: center;
		gap: 2px;
		padding: 2px;
		border-radius: var(--pptx-radius, var(--office-radius, 4px));
		background: var(--pptx-muted, var(--office-surface, #f3f4f6));
	}

	.pptx-svelte-inspector-tabs button {
		padding: 3px 8px;
		border: none;
		border-radius: calc(var(--pptx-radius, var(--office-radius, 4px)) - 2px);
		background: transparent;
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
		cursor: pointer;
		font: inherit;
		font-size: 11px;
		white-space: nowrap;
	}

	.pptx-svelte-inspector-tabs button:hover {
		background: var(--pptx-accent, var(--office-selected, #e5e7eb));
		color: var(--pptx-accent-foreground, var(--office-foreground, #1f2937));
	}

	.pptx-svelte-inspector-tab-active {
		background: var(--pptx-primary, var(--office-accent, #2563eb)) !important;
		color: #fff !important;
	}

	.pptx-svelte-inspector-body {
		flex: 1 1 auto;
		min-height: 0;
		overflow-y: auto;
		padding: 0 12px 12px;
	}

	.pptx-svelte-inspector-body > :global(.pptx-svelte-comments) {
		width: 100%;
		padding-left: 0;
		border-left: none;
		margin-top: 12px;
	}

	.pptx-svelte-inspector-body > :global(.pptx-svelte-layers) {
		margin-top: 12px;
	}

	.pptx-svelte-inspector-empty {
		margin: 12px 0 0;
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
	}

	/*
	 * Touch target below MOBILE_BREAKPOINT (768, matches
	 * pptx-viewer-shared's isDensePanelCompact/MIN_TOUCH_TARGET_PX): every
	 * sub-panel this inspector hosts (Fill/Stroke, Text, Animation, Chart,
	 * Table, ...) is its own Svelte component with its own scoped
	 * `select,button{height:26px}`-style desktop-mouse sizing (ShapeSection,
	 * FillStrokeSection, TextSection, etc.). Rather than hand-editing every
	 * one of those sub-panels' scoped styles (CLAUDE.md Rule 2: fix the
	 * cross-cutting concern once), `:global()` reaches through the scoping
	 * boundary from this shared host so every nested control clears the WCAG
	 * target from a single place.
	 *
	 * `!important` on the button rule is load-bearing:
	 * `ViewerGlobalStyles.svelte`'s document-level
	 * `:global(.pptx-svelte-viewer :is(button, [role='button']):not(...):not(...))`
	 * baseline reset out-specificities this rule's `[attr][hash] button`
	 * (0,2,1) with its own chained :not() pseudo-classes (0,4,0), regardless
	 * of source order. `select` is not matched by that button-only baseline,
	 * so it does not need `!important`.
	 */
	@media (max-width: 767px) {
		[data-pptx-inspector] :global(button) {
			min-width: 44px !important;
			min-height: 44px !important;
		}
		[data-pptx-inspector] :global(select) {
			min-height: 44px;
		}
		[data-pptx-inspector] :global(pptx-ui-select) {
			min-height: 44px;
		}
	}
	[data-pptx-inspector] :global(pptx-ui-select) {
		width: 100%;
		min-width: 0;
	}
</style>
