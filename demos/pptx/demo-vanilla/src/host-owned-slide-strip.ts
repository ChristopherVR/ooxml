import { disposePptxRenderStage, renderSlideThumbnail } from 'pptx-vanilla-viewer';
import type { LoadedPresentation, Store, ViewerState } from 'pptx-vanilla-viewer';

import {
	cssText,
	HOST_STRIP_LABEL,
	HOST_STRIP_STYLE,
	HOST_STRIP_THUMBNAIL_WIDTH,
	hostStripItemLabel,
	hostStripItemStyle,
} from '../../shared/host-thumbnail-strip';
import { t } from './demo-i18n';

export interface HostSlideStripOptions {
	store: Store<ViewerState>;
	getLoaded(): LoadedPresentation | undefined;
	/** Commit an open inline text edit first; false keeps the editor open (invalid edit). */
	commitInline(): boolean;
}

/**
 * A host-owned thumbnail strip (`?hostStrip=1`) in place of `createSlidesPane`:
 * one `renderSlideThumbnail` per slide, redrawn from the shell's store whenever
 * the slides change. Each redraw retires the previous thumbnails' render
 * subscriptions first, as the thumbnail's cleanup contract asks.
 */
export function mountHostSlideStrip(
	container: HTMLElement,
	options: HostSlideStripOptions,
): { destroy(): void } {
	const { store } = options;
	const nav = document.createElement('nav');
	nav.setAttribute('aria-label', HOST_STRIP_LABEL);
	nav.dataset.hostSlideStrip = '';
	nav.style.cssText = cssText(HOST_STRIP_STYLE);
	container.append(nav);
	const draw = (): void => {
		const state = store.get();
		const loaded = options.getLoaded();
		disposePptxRenderStage(nav);
		nav.replaceChildren(
			...state.slides.map((slide, index) => {
				const button = document.createElement('button');
				button.type = 'button';
				button.setAttribute('aria-label', hostStripItemLabel(index));
				if (index === state.currentSlide) {
					button.setAttribute('aria-current', 'true');
				}
				button.style.cssText = cssText(hostStripItemStyle(index === state.currentSlide));
				button.onclick = () => {
					if (index !== store.get().currentSlide && options.commitInline()) {
						store.set({ currentSlide: index, selectedElementId: null, selectedElementIds: [] });
					}
				};
				button.append(
					renderSlideThumbnail({
						document,
						slide,
						canvasSize: state.canvasSize,
						previewWidth: HOST_STRIP_THUMBNAIL_WIDTH,
						mediaDataUrls: loaded?.mediaDataUrls,
						headerFooter: loaded?.headerFooter,
						customProperties: loaded?.customProperties,
						firstSlideNumber: loaded?.firstSlideNumber,
						colorScheme: loaded?.colorScheme,
						fontScheme: loaded?.fontScheme,
						formatScheme: loaded?.formatScheme,
						tableStyleMap: loaded?.tableStyleMap,
						slides: state.slides,
						t,
					}),
				);
				return button;
			}),
		);
	};
	const unsubscribe = store.subscribe((state, previous) => {
		if (
			state.slides !== previous.slides ||
			state.currentSlide !== previous.currentSlide ||
			state.canvasSize !== previous.canvasSize
		) {
			draw();
		}
	});
	draw();
	return {
		destroy() {
			unsubscribe();
			disposePptxRenderStage(nav);
			nav.remove();
		},
	};
}
