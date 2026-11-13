<script lang="ts">
	/**
	 * BuildingBlocksCanvas: the slide canvas block. The scaled slide stage with the editing layer
	 * over the viewer state `createViewerBuildingBlocks` made, in a scrollable viewport that reports
	 * its size back (the fit scale reads it). It mounts no ribbon, thumbnails, status bar, notes or
	 * dialogs: gestures that need a dialog raise a building-block request instead.
	 */
	import { markFindHighlights, themeToCssVars } from 'ooxml-ui/pptx';
	import { tick } from 'svelte';
	import type { ViewerTheme } from '../viewer/types';

	import CollaborationCursors from '../viewer/collab/components/CollaborationCursors.svelte';
	import RemoteSelectionOverlay from '../viewer/collab/components/RemoteSelectionOverlay.svelte';
	import EditorChromeStyle from '../viewer/components/EditorChromeStyle.svelte';
	import EditorLayer from '../viewer/components/EditorLayer.svelte';
	import SlideCanvas from '../viewer/components/SlideCanvas.svelte';
	import { createEditCommits } from '../viewer/components/viewer-body-commits';
	import ViewerGlobalStyles from '../viewer/components/ViewerGlobalStyles.svelte';
	import { styleToString } from '../viewer/style';
	import { viewportClientSize } from '../viewer/viewport-client-size';
	import type { ViewerBuildingBlocks } from './types';

	const { blocks, theme, class: className = '' }: {
		blocks: ViewerBuildingBlocks;
		/** A viewer theme, as `PowerPointViewer`'s `theme`; unset inherits the host's tokens. */
		theme?: ViewerTheme | undefined;
		class?: string;
	} = $props();

	const vs = $derived(blocks.viewerState);
	const commits = $derived(createEditCommits(vs.editor));
	const rootStyle = $derived(theme ? styleToString(themeToCssVars(theme)) : undefined);
	let holder = $state.raw<HTMLDivElement>();

	function register(node: HTMLDivElement) {
		blocks.canvas.setRoot(node);
		return { destroy: () => blocks.canvas.setRoot(undefined) };
	}

	// Outline the elements holding the host find bar's matches (the current one heavier).
	$effect(() => {
		const targets = blocks.findHighlights;
		void vs.activeSlide;
		if (!holder || targets.length === 0) {
			return;
		}
		const root = holder;
		let stale = false;
		void tick().then(() => {
			if (!stale) {
				markFindHighlights(root, targets);
			}
			return undefined;
		});
		return () => {
			stale = true;
			markFindHighlights(root, []);
		};
	});
</script>

<!-- The canvas is keyboard-focusable for native selection and arrow-key commands. -->
<!-- svelte-ignore a11y_no_noninteractive_tabindex, a11y_no_noninteractive_element_interactions -->
<div
	use:register
	class={`pptx-svelte-viewer pptx-svelte-blocks-canvas ${className}`}
	style={rootStyle}
	data-pptx-block="canvas"
	data-pptx-theme-root
	data-pptx-editor-chrome
	role="application"
	aria-label={blocks.t('pptx.titleBar.defaultFileName')}
	aria-busy={vs.loader.loading}
	tabindex="0"
	onkeydown={vs.onKeydown}
>
	<EditorChromeStyle />
	<ViewerGlobalStyles />
	<div class="pptx-svelte-blocks-viewport" data-pptx-viewport use:viewportClientSize={blocks.canvas.setViewportSize}>
		{#if vs.loader.error}<p role="alert">{vs.loader.error}</p>{/if}
		<SlideCanvas
			slide={vs.activeSlide}
			canvasSize={vs.loader.canvasSize}
			mediaDataUrls={vs.loader.mediaDataUrls}
			scale={vs.scale}
			editingActive={vs.editingActive}
			editTemplateMode={vs.editor.editTemplateMode}
			editingElementId={vs.controller.editingId}
			selectedElementIds={vs.editor.selection.ids}
			ontablecellcommit={vs.editingActive ? commits.commitTableCell : undefined}
			onsmartartnodecommit={vs.editingActive ? commits.commitSmartArtNode : undefined}
			onsmartartnodefill={vs.editingActive ? commits.commitSmartArtFill : undefined}
			onchartpointcommit={vs.editingActive ? commits.commitChartPoint : undefined}
			ontableresizecolumns={vs.editingActive ? commits.commitTableResizeColumns : undefined}
			ontableresizerow={vs.editingActive ? commits.commitTableResizeRow : undefined}
			onstageholder={(element) => {
				holder = element ?? undefined;
				blocks.canvas.setStageHolder(element ?? undefined);
			}}
			onstagepointerdown={vs.controller.onStagePointerDown}
			onstagepointermove={vs.controller.onStagePointerMove}
			onstagedblclick={vs.controller.onStageDblClick}
			onstageclick={vs.controller.onStageClick}
		>
			{#if vs.editingActive}<EditorLayer controller={vs.controller} scale={vs.scale} />{/if}
			<RemoteSelectionOverlay
				presences={vs.collab.remotePresences}
				elements={vs.activeSlide?.elements ?? []}
				activeSlideIndex={vs.viewer.current}
				zoom={vs.scale}
			/>
			<CollaborationCursors cursors={vs.collab.cursors} zoom={vs.scale} />
		</SlideCanvas>
	</div>
</div>

<style>
	.pptx-svelte-blocks-canvas {
		position: relative;
		display: flex;
		flex-direction: column;
		width: 100%;
		height: 100%;
		min-width: 0;
		min-height: 0;
		outline: none;
	}

	.pptx-svelte-blocks-viewport {
		display: flex;
		flex: 1;
		min-width: 0;
		min-height: 0;
		overflow: auto;
		padding: 16px;
		box-sizing: border-box;
	}

	/* The elements holding the host find bar's matches; reapplied by the effect above. */
	:global(.pptx-svelte-blocks-canvas [data-pptx-find-match]) {
		outline: 2px dashed var(--pptx-primary, var(--office-primary, #2563eb));
		outline-offset: 2px;
	}

	:global(.pptx-svelte-blocks-canvas [data-pptx-find-match='current']) {
		outline-style: solid;
		outline-width: 3px;
	}
</style>
