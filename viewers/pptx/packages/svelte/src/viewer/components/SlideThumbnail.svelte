<script lang="ts">
	/**
	 * SlideThumbnail: one slide drawn the way `SlidesPane` draws it, for a host
	 * that keeps its own thumbnail strip.
	 *
	 * It paints the slide's template (layout/master) elements beneath its own,
	 * with the slide background, through the same static `SlideStage` the rail
	 * uses, scaled to `previewWidth`. The stage is static: no selection,
	 * editing, pointer handlers, focus stops or element ids, so it never
	 * competes with the editing canvas. Fields and table styles resolve from the
	 * same props `SlidesPane` takes, so a host feeds both from one place.
	 */
	import { buildPreviewElements, EDITOR_THUMBNAIL_WIDTH, slideThumbnailLayout } from 'ooxml-ui/pptx';
	import type { PptxSlide } from 'pptx-viewer-core';

	import { provideSlidesPaneContexts } from './slides-pane-contexts';
	import type { SlideThumbnailProps } from './slide-thumbnail-props';
	import SlideStage from './SlideStage.svelte';

	const props: SlideThumbnailProps = $props();

	provideSlidesPaneContexts(() => props);

	const EMPTY_MEDIA = new Map<string, string>();

	const stageSize = $derived({
		width: Math.max(1, props.canvasSize.width),
		height: Math.max(1, props.canvasSize.height),
	});
	const layout = $derived(
		slideThumbnailLayout(
			stageSize.width,
			stageSize.height,
			props.previewWidth ?? EDITOR_THUMBNAIL_WIDTH,
		),
	);
	// Template first, slide elements on top; `buildPreviewElements` also honours
	// "Hide Background Graphics".
	const previewSlide = $derived<PptxSlide>({
		...props.slide,
		elements: buildPreviewElements(props.slide, props.templateElements ?? []),
	});

	/**
	 * `SlideStage` keeps the `data-element-id` markers on static stages in this
	 * binding, but they must stay unique to the editing canvas: hit-testing,
	 * animation preview and element queries take the first match. Strip them
	 * after every render, and from nodes a renderer adds later (async charts,
	 * media), so a host strip never shadows the canvas copy.
	 */
	function stripMarkers(root: ParentNode | undefined): void {
		for (const marked of root?.querySelectorAll('[data-element-id]') ?? []) {
			marked.removeAttribute('data-element-id');
		}
	}
	function staticMarkers(node: HTMLElement) {
		stripMarkers(node);
		const observer = new MutationObserver(() => stripMarkers(node));
		observer.observe(node, {
			subtree: true,
			childList: true,
			attributes: true,
			attributeFilter: ['data-element-id'],
		});
		return { destroy: () => observer.disconnect() };
	}

	// Re-strip in the same flush as a re-render (template effects run before
	// this one), so the markers never show between an edit and the observer.
	let box: HTMLElement | undefined = $state();
	$effect(() => {
		void previewSlide;
		stripMarkers(box);
	});
</script>

<div
	bind:this={box}
	use:staticMarkers
	class="pptx-svelte-slide-thumbnail"
	style={`position: relative; width: 100%; height: ${layout.height}px; overflow: hidden`}
>
	<SlideStage
		slide={previewSlide}
		canvasSize={stageSize}
		mediaDataUrls={props.mediaDataUrls ?? EMPTY_MEDIA}
		scale={layout.scale}
	/>
</div>
