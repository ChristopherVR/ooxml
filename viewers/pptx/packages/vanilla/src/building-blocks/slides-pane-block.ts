import type { PptxViewer } from '../viewer/PptxViewer';
import { createSlidesPane } from '../viewer/ui/slides-pane';
import type { SlidesPane, SlidesPaneOptions } from '../viewer/ui/slides-pane';
import type { ViewerBuildingBlocksOptions } from './types';

/** The rail's options from the viewer's store: the same deck, active slide and slide commands. */
function paneOptions(
	viewer: PptxViewer | null,
	options: ViewerBuildingBlocksOptions,
): SlidesPaneOptions {
	if (!viewer) {
		return {
			slides: [],
			activeSlideIndex: 0,
			canvasSize: { width: 960, height: 540 },
			onSelectSlide: () => {},
		};
	}
	const state = viewer.store.get();
	const pane: SlidesPaneOptions = {
		slides: state.slides,
		templateElementsBySlideId: state.templateElementsBySlideId,
		activeSlideIndex: state.currentSlide,
		canvasSize: state.canvasSize,
		mediaDataUrls: state.mediaDataUrls,
		sections: state.sections,
		canEdit: state.editable,
		headerFooter: state.headerFooter,
		customProperties: state.customProperties,
		t: viewer.t,
		onElementRenderError: (error, context) => {
			const report = options.onElementRenderError;
			if (report) {
				report(error, context);
			} else {
				console.error(error, context);
			}
		},
		onSelectSlide: (index) => viewer.goToSlide(index),
		onAddSlide: () => viewer.editor.getEditActions().addSlide(),
		onDuplicateSlides: (indexes) => viewer.duplicateSlides(indexes),
		onDeleteSlides: (indexes) => viewer.deleteSlides(indexes),
		onHideSlides: (indexes) => viewer.toggleHideSlides(indexes),
	};
	if (state.firstSlideNumber !== undefined) {
		pane.firstSlideNumber = state.firstSlideNumber;
	}
	if (state.colorScheme) {
		pane.colorScheme = state.colorScheme;
	}
	if (state.fontScheme) {
		pane.fontScheme = state.fontScheme;
	}
	if (state.formatScheme) {
		pane.formatScheme = state.formatScheme;
	}
	if (state.tableStyleMap) {
		pane.tableStyleMap = state.tableStyleMap;
	}
	if (options.onOpenLayoutGallery) {
		const open = options.onOpenLayoutGallery;
		pane.onOpenLayoutGallery = (request) => open(request);
	}
	return pane;
}

/** Mount the rail into `host` and keep it in step with the viewer through `subscribe`. */
export function mountSlidesPaneBlock(
	host: HTMLElement,
	read: () => { viewer: PptxViewer | null; options: ViewerBuildingBlocksOptions },
	subscribe: (listener: () => void) => () => void,
): () => void {
	const first = read();
	const pane: SlidesPane = createSlidesPane(host, paneOptions(first.viewer, first.options));
	const stop = subscribe(() => {
		const { viewer, options } = read();
		pane.update(paneOptions(viewer, options));
	});
	return () => {
		stop();
		pane.destroy();
	};
}
