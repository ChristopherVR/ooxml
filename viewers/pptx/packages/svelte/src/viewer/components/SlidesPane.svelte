<script lang="ts">
	/**
	 * SlidesPane: the slide thumbnail rail as a standalone building block, for a
	 * host that composes it beside its own `SlideCanvas`. It renders the same
	 * `ThumbnailRail` as `PowerPointViewer`, behind a flat, host-driven contract
	 * (`SlidesPaneProps`); `createSlidesPaneProps` builds it from a
	 * `createViewerState` bag.
	 *
	 * Thumbnails read the field and table style contexts through Svelte
	 * context. This component re-provides both from its props, so they match the
	 * canvas without an enclosing `PowerPointViewer`; an omitted prop falls back
	 * to the enclosing `createViewerState` context.
	 */
	import { openSlidesPaneLayoutGallery } from 'ooxml-ui/pptx';

	import { getFieldContextGetter, provideFieldContext } from '../state/field-context';
	import { getRenderContextSource, provideRenderContext } from '../state/render-context';
	import type { SlidesPaneProps } from './slides-pane-props';
	import ThumbnailRail from './ThumbnailRail.svelte';

	const props: SlidesPaneProps = $props();

	// Svelte context is captured once at init, so the provided values are
	// getters over the live props, falling back to the enclosing context.
	const inheritedFieldContext = getFieldContextGetter();
	provideFieldContext(() => props.fieldContext ?? inheritedFieldContext?.());
	const inheritedRenderContext = getRenderContextSource();
	provideRenderContext({
		getColorScheme: () =>
			props.tableStyleContext
				? props.tableStyleContext.theme?.colorScheme
				: inheritedRenderContext?.getColorScheme(),
		getTableStyleMap: () =>
			props.tableStyleContext
				? props.tableStyleContext.tableStyleMap
				: inheritedRenderContext?.getTableStyleMap(),
		getFontScheme: () =>
			props.tableStyleContext
				? props.tableStyleContext.theme?.fontScheme
				: inheritedRenderContext?.getFontScheme?.(),
		getFormatScheme: () =>
			props.tableStyleContext
				? props.tableStyleContext.theme?.formatScheme
				: inheritedRenderContext?.getFormatScheme?.(),
		getCanvasSize: () => props.canvasSize,
	});

	function openLayoutForSlide(slideIndex: number, x: number, y: number): void {
		openSlidesPaneLayoutGallery({ slideIndex, x, y }, props.onselectslide, props.onopenlayoutgallery);
	}
</script>

<ThumbnailRail
	slides={props.slides}
	sections={props.sections}
	canvasSize={props.canvasSize}
	mediaDataUrls={props.mediaDataUrls}
	current={props.activeSlideIndex}
	onselect={props.onselectslide}
	editable={props.canEdit ?? false}
	onmove={props.onmoveslide}
	onaddslide={props.onaddslide}
	onaddslideafter={props.onaddslideafter}
	onduplicateslides={props.onduplicateslides}
	ondeleteslides={props.ondeleteslides}
	ontogglehideslides={props.onhideslides}
	onopenlayoutforslide={openLayoutForSlide}
	onaddsectionat={props.onaddsection}
	onsectiontoggle={props.ontogglesection}
	onsectionrename={props.onrenamesection}
	onsectiondelete={props.ondeletesection}
	onsectionmove={(id, direction) =>
		direction === 'up' ? props.onmovesectionup?.(id) : props.onmovesectiondown?.(id)}
/>
