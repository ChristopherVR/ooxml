<script lang="ts">
	/** A host-owned thumbnail strip, fed from the props the built-in slides pane takes. */
	import { SlideThumbnail } from 'pptx-svelte-viewer';
	import type { SlidesPaneProps } from 'pptx-svelte-viewer';

	import {
		cssText,
		HOST_STRIP_LABEL,
		HOST_STRIP_STYLE,
		HOST_STRIP_THUMBNAIL_WIDTH,
		hostStripItemLabel,
		hostStripItemStyle,
	} from '../../shared/host-thumbnail-strip';

	const { pane }: { pane: SlidesPaneProps } = $props();
</script>

<nav aria-label={HOST_STRIP_LABEL} data-host-slide-strip style={cssText(HOST_STRIP_STYLE)}>
	{#each pane.slides as slide, index (slide.id)}
		<button
			type="button"
			aria-label={hostStripItemLabel(index)}
			aria-current={index === pane.activeSlideIndex ? 'true' : undefined}
			style={cssText(hostStripItemStyle(index === pane.activeSlideIndex))}
			onclick={() => pane.onselectslide(index)}
		>
			<SlideThumbnail
				{slide}
				canvasSize={pane.canvasSize}
				previewWidth={HOST_STRIP_THUMBNAIL_WIDTH}
				mediaDataUrls={pane.mediaDataUrls}
				fieldContext={pane.fieldContext}
				tableStyleContext={pane.tableStyleContext}
			/>
		</button>
	{/each}
</nav>
