<script lang="ts">
	/**
	 * ViewerBody: the viewer's three-column layout, thumbnail rail | slide
	 * viewport + notes panel | inspector. Split out of `PowerPointViewer.svelte`
	 * to keep that file under the repo's file-size budget; purely presentational,
	 * all state/logic stays owned by the parent.
	 *
	 * Everything inside the viewport (placeholders, the scaled stage, its ruler
	 * strips and overlay stack, the context menu) lives in `ViewerStage`, and the
	 * prop contract lives in `viewer-body-props.ts`, so this file stays layout
	 * only: it measures the viewport and decides which of the three columns are
	 * present, nothing more.
	 */
	import type { PptxLayoutOption, PptxLayoutPreview } from 'pptx-viewer-core';
	import { EDITOR_SLIDE_RAIL_WIDTH, MAX_ZOOM_SCALE, MIN_ZOOM_SCALE, observeElementHeight } from 'ooxml-ui/pptx';

	import { canvasPinchZoom } from '../canvas-pinch-zoom';
	import { viewportClientSize } from '../viewport-client-size';
	import InspectorPanel from './inspector/InspectorPanel.svelte';
	import NotesPanel from './NotesPanel.svelte';
	import LayoutGalleryMenu from './ribbon/home/LayoutGalleryMenu.svelte';
	import ThumbnailRail from './ThumbnailRail.svelte';
	import ViewerRegion from './ViewerRegion.svelte';
	import { useNotesBarHeightSink } from '../state/notes-bar-height.svelte';
	import { useViewerCustomization } from '../state/viewer-customization.svelte';
	import type { ViewerBodyProps } from './viewer-body-props';
	import ViewerStage from './ViewerStage.svelte';

	const {
		t,
		editor,
		handler,
		presentationTheme,
		onthemechange,
		chromeVisible,
		showThumbnails,
		showNotes,
		displaySlides,
		canvasSize,
		mediaDataUrls,
		current,
		onselect,
		loading,
		isEncrypted,
		error,
		activeSlide,
		scale,
		presenting,
		gridSpacingPx,
		presentationTransition,
		onTransitionDone,
		onAdvance,
		onPresentationContextMenu,
		editingActive,
		blackout = 'none',
		controller,
		onstageresize,
		onstageholder,
		getzoomscale,
		onpinchzoom,
		notesExpanded,
		onNotesCommit,
		onNotesToggle,
		collabCursors = [],
		collabPresences = [],
		contextMenu,
		onContextMenuClose,
		canvasContextMenu,
		onCanvasContextMenuClose,
		parityUi,
		onmoveSlide,
		annotations,
		guides = [],
		onchangeguide,
		ondeleteguide,
		onaddguide,
		showRulers = false,
		rulerUnit = 'inches',
		spellCheck = false,
		chromeUi,
		aiPickMode = false,
		aiActive = false,
		aiHighlights = [],
		aiChangeBatch = null,
		onaipickelement,
		onaskai,
		onfixai,
	}: ViewerBodyProps = $props();

	// The fractional client size of the viewport, written by the `viewportClientSize` action.
	let viewportWidth = $state(0);
	let viewportHeight = $state(0);
	function setViewportSize(size: { width: number; height: number }): void {
		viewportWidth = size.width;
		viewportHeight = size.height;
	}
	$effect(() => {
		onstageresize(viewportWidth, viewportHeight);
	});

	// The thumbnail rail's "Layout" command opens the same gallery the ribbon's
	// Home > Layout button does, anchored at the click point.
	let railLayoutGalleryAnchor = $state<{ x: number; y: number } | null>(null);
	let railLayoutOptions = $state<PptxLayoutOption[]>([]);
	let railLayoutPreviews = $state<ReadonlyMap<string, PptxLayoutPreview>>(new Map());
	function openLayoutForSlide(index: number, x: number, y: number): void {
		onselect(index);
		railLayoutGalleryAnchor = { x, y };
		void editor.slidesOps.availableLayouts().then((options) => (railLayoutOptions = options));
		void editor.slidesOps
			.layoutPreviews()
			.then((previews) => (railLayoutPreviews = previews));
	}
	// Host-removed regions (`hiddenPanels`): slides pane, notes, inspector.
	const customization = useViewerCustomization();
	// A replaced notes panel keeps the slot (content-driven height, collapse via display:none) and
	// reports its live height like the built-in strip, so the compat toasts clear it.
	const notesBarSink = useNotesBarHeightSink();
	function reportNotesHeight(node: HTMLElement, expanded: boolean) {
		let stop: (() => void) | undefined;
		const start = (): void => {
			stop?.();
			stop = undefined;
			if (expanded && notesBarSink) {
				stop = observeElementHeight(node, (height) => notesBarSink.report(height));
			} else {
				notesBarSink?.report(0);
			}
		};
		start();
		return {
			update(next: boolean): void {
				expanded = next;
				start();
			},
			destroy(): void {
				stop?.();
				notesBarSink?.report(0);
			},
		};
	}
</script>

<div class="pptx-svelte-body-stack">
<div class="pptx-svelte-body" data-pptx-chrome="body">
	{#if showThumbnails && !chromeUi?.sidebarCollapsed && chromeVisible && displaySlides.length > 0 && customization.isPanelVisible('slidesPane')}
		{@const slidesRenderer = customization.regionRenderer('slidesPane')}
		{#if slidesRenderer}
			<nav data-pptx-chrome="slides" class="pptx-svelte-slides-region" style={`width:${EDITOR_SLIDE_RAIL_WIDTH}px`} aria-label={t('pptx.sections.slides')}>
				<ViewerRegion region="slidesPane" renderer={slidesRenderer} />
			</nav>
		{:else}
		<ThumbnailRail
			slides={displaySlides}
			sections={editor.sections}
			{canvasSize}
			{mediaDataUrls}
			{current}
			{onselect}
			editable={editingActive}
			onmove={onmoveSlide}
			onaddslide={() => {
				const index = editor.slidesOps.insertSlideAfterCurrent();
				if (index !== null) onselect(index);
			}}
			onaddslideafter={(index) => {
				const newIndex = editor.slidesOps.insertSlideAfter(index);
				if (newIndex !== null) onselect(newIndex);
			}}
			onduplicateslides={(indexes) => editor.slidesOps.duplicateSlides(indexes)}
			ondeleteslides={(indexes) => editor.slidesOps.deleteSlides(indexes)}
			ontogglehideslides={(indexes) => editor.slidesOps.toggleSlidesHidden(indexes)}
			onopenlayoutforslide={openLayoutForSlide}
			onaddsectionat={(index) => editor.sectionOps.add(t('pptx.sections.defaultName'), index)}
			onsectiontoggle={(id) => editor.sectionOps.toggle(id)}
			onsectionrename={(id, name) => editor.sectionOps.rename(id, name)}
			onsectiondelete={(id) => editor.sectionOps.delete(id)}
			onsectionmove={(id, direction) =>
				direction === 'up' ? editor.sectionOps.moveUp(id) : editor.sectionOps.moveDown(id)}
		/>
		{/if}
	{/if}
	{#if railLayoutGalleryAnchor}
		<!-- svelte-ignore a11y_click_events_have_key_events -->
		<div
			class="pptx-svelte-context-backdrop"
			aria-hidden="true"
			onclick={() => (railLayoutGalleryAnchor = null)}
			oncontextmenu={(event) => { event.preventDefault(); railLayoutGalleryAnchor = null; }}
		></div>
		<div
			class="pptx-svelte-layout-gallery-anchor"
			style={`left:${railLayoutGalleryAnchor.x}px;top:${railLayoutGalleryAnchor.y}px`}
		>
			<LayoutGalleryMenu
				layouts={railLayoutOptions}
				previews={railLayoutPreviews}
				currentLayoutPath={editor.slides[current]?.layoutPath}
				onselect={(layout) => {
					void editor.slidesOps.applyLayout(layout.path);
					railLayoutGalleryAnchor = null;
				}}
			/>
		</div>
	{/if}
	<div class="pptx-svelte-main">
		<div
			class="pptx-svelte-viewport"
			use:viewportClientSize={setViewportSize}
			data-pptx-viewport
			use:canvasPinchZoom={{
				getScale: getzoomscale,
				minScale: MIN_ZOOM_SCALE,
				maxScale: MAX_ZOOM_SCALE,
				onPinchZoom: onpinchzoom,
			}}
		>
			<ViewerStage
				{t}
				{editor}
				{controller}
				{canvasSize}
				{mediaDataUrls}
				{current}
				{loading}
				{isEncrypted}
				{error}
				{activeSlide}
				{scale}
				{presenting}
				{gridSpacingPx}
				{presentationTransition}
				{onTransitionDone}
				{onAdvance}
				{onPresentationContextMenu}
				{editingActive}
				{blackout}
				{onstageholder}
				{collabCursors}
				{collabPresences}
				{contextMenu}
				{onContextMenuClose}
				{canvasContextMenu}
				{onCanvasContextMenuClose}
				{parityUi}
				{annotations}
				{guides}
				{onchangeguide}
				{ondeleteguide}
				{onaddguide}
				{showRulers}
				{rulerUnit}
				{spellCheck}
				{chromeUi}
				{aiPickMode}
				{aiActive}
				{aiHighlights}
				{aiChangeBatch}
				{onaipickelement}
				{onaskai}
				{onfixai}
			/>
		</div>

	</div>
	{#if editingActive && chromeVisible && displaySlides.length > 0 && chromeUi?.inspectorOpen !== false && customization.isPanelVisible('inspector')}
		<InspectorPanel {editor} {handler} {presentationTheme} {onthemechange} {mediaDataUrls} ui={chromeUi} {canvasSize} />
	{/if}
</div>
		{#if showNotes && chromeVisible && displaySlides.length > 0 && customization.isPanelVisible('notes')}
			{@const notesRenderer = customization.regionRenderer('notes')}
			{#if notesRenderer}
				<section class="pptx-svelte-notes-region" data-pptx-chrome="notes" data-collapsed={!notesExpanded} hidden={!notesExpanded} use:reportNotesHeight={notesExpanded}>
					<ViewerRegion region="notes" renderer={notesRenderer} />
				</section>
			{:else}
			<NotesPanel
				slide={activeSlide}
				expanded={notesExpanded}
				onupdate={onNotesCommit}
				ontoggle={onNotesToggle}
				notesStyle={editor.notesMaster?.notesStyle}
			/>
			{/if}
		{/if}
</div>

<style>
	.pptx-svelte-body-stack { display: flex; flex-direction: column; flex: 1; min-height: 0; }
	.pptx-svelte-body {
		display: flex;
		flex: 1;
		min-height: 0;
	}

	.pptx-svelte-slides-region { display: flex; flex: none; flex-direction: column; min-height: 0; overflow: hidden; }
	.pptx-svelte-notes-region { display: flex; flex: none; flex-direction: column; min-height: 0; overflow: hidden; }
	.pptx-svelte-notes-region[hidden] { display: none; }

	.pptx-svelte-main {
		display: flex;
		flex-direction: column;
		flex: 1;
		min-width: 0;
		min-height: 0;
	}

	/* Thumbnail rail's Layout gallery: a zero-size anchor at the click point,
	   plus a click-outside backdrop, mirroring `ViewerStage`'s canvas-menu one. */
	.pptx-svelte-context-backdrop {
		position: fixed;
		inset: 0;
		z-index: 119;
	}
	.pptx-svelte-layout-gallery-anchor {
		position: fixed;
		z-index: 120;
	}

	.pptx-svelte-viewport {
		flex: 1;
		display: flex;
		overflow: auto;
		min-width: 0;
		min-height: 0;
	}
</style>
