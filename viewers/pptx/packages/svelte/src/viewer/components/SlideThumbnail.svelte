<script lang="ts">
	/**
	 * SlideThumbnail: one slide drawn the way `SlidesPane` draws it, for a host
	 * that keeps its own thumbnail strip.
	 *
	 * It paints the slide's template (layout/master) elements beneath its own,
	 * with the slide background, through the same static `SlideStage` the rail
	 * uses, scaled to `previewWidth`. The stage is static: no selection,
	 * editing, pointer handlers, focus stops or element ids (a static
	 * `SlideStage` drops them, as the rail's do), so it never competes with the
	 * editing canvas. Fields and table styles resolve from the
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
</script>

<div
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
