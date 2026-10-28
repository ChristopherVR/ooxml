<script lang="ts">
	/**
	 * Test-only custom shell for `SlidesPane`: builds `createViewerState` and
	 * renders the rail beside a `SlideCanvas`, the way a host's own shell does
	 * (see the custom-shell demo). Hands the state bag back through `onready`.
	 */
	import type { SlidesPaneLayoutGalleryRequest } from 'ooxml-ui/pptx';
	import { onDestroy } from 'svelte';

	import { createTranslator } from '../../i18n/translator';
	import { createViewerState } from '../state/create-viewer-state.svelte';
	import type { ViewerStateBag } from '../state/create-viewer-state-types';
	import SlideCanvas from './SlideCanvas.svelte';
	import { createSlidesPaneProps } from './slides-pane-props';
	import SlidesPane from './SlidesPane.svelte';

	const {
		source,
		editable = false,
		onready,
		onopenlayoutgallery,
		ondirtychange,
	}: {
		source: Uint8Array;
		editable?: boolean;
		onready: (state: ViewerStateBag) => void;
		onopenlayoutgallery?: (request: SlidesPaneLayoutGalleryRequest) => void;
		ondirtychange?: (dirty: boolean) => void;
	} = $props();

	const state = createViewerState({
		getSource: () => source,
		getAutosave: () => false,
		getFilePath: () => undefined,
		getInitialSlide: () => 0,
		t: createTranslator(() => 'en'),
		getSmartArt3D: () => false,
		getSurfaceChart3D: () => false,
		getBarChart3D: () => false,
		getLineChart3D: () => false,
		getAreaChart3D: () => false,
		getPieChart3D: () => false,
		getEditable: () => editable,
		ondirtychange: (dirty) => ondirtychange?.(dirty),
		getStageHolderEl: () => undefined,
		getRootEl: () => undefined,
		getViewportWidth: () => 0,
		getViewportHeight: () => 0,
		getMasterScale: () => 1,
	});
	onDestroy(() => state.destroy());
	const paneProps = createSlidesPaneProps(state, {
		onopenlayoutgallery: (request) => onopenlayoutgallery?.(request),
	});

	// svelte-ignore state_referenced_locally -- a one-time synchronous hand-off.
	onready(state);
</script>

<div data-test-shell>
	<div data-test-rail><SlidesPane {...paneProps} /></div>
	<div data-test-canvas>
		<SlideCanvas
			slide={state.activeSlide}
			canvasSize={state.loader.canvasSize}
			mediaDataUrls={state.loader.mediaDataUrls}
			scale={1}
		/>
	</div>
</div>
