import {
	createFindBarController,
	createPptxBuildingBlockSession,
	describeBuildingBlockToolbar,
	runBuildingBlockToolbarControl,
} from 'ooxml-ui/pptx';
import type { ViewerMode } from 'ooxml-ui/pptx';

import { createTranslator } from '../viewer/i18n';
import type { PptxViewerSource } from '../viewer/load';
import { clampSlideIndex as clampIndex } from '../viewer/state';
import type { PptxViewerOptions } from '../viewer/types';
import { BuildingBlocksViewer } from './blocks-viewer';
import { createVanillaPort } from './session-port';
import { mountSlidesPaneBlock } from './slides-pane-block';
import { createToolbarView } from './toolbar-view';
import type {
	ViewerBuildingBlocks,
	ViewerBuildingBlocksOptions,
	ViewerBuildingBlocksState,
} from './types';

/** The regions the canvas block leaves out: the host composes its own shell around it. */
const BLOCK_HIDDEN_PANELS = ['statusBar', 'inspector', 'notes', 'titleBar', 'quickAccessToolbar'];

function viewerOptions(options: ViewerBuildingBlocksOptions): PptxViewerOptions {
	// The building-block-only options never reach the viewer.
	const {
		content,
		canEdit,
		onRequest: _onRequest,
		fileLimits: _fileLimits,
		onOpenLayoutGallery: _onOpenLayoutGallery,
		...rest
	} = options;
	const hiddenPanels = [...(rest.customization?.hiddenPanels ?? []), ...BLOCK_HIDDEN_PANELS];
	return {
		...rest,
		...(content ? { source: content } : {}),
		editable: canEdit === true,
		showToolbar: false,
		showThumbnails: false,
		showCompatibilityToasts: false,
		customization: { ...rest.customization, hiddenPanels } as PptxViewerOptions['customization'],
	};
}

async function load(viewer: BuildingBlocksViewer, source: PptxViewerSource): Promise<void> {
	await (typeof source === 'string' ? viewer.loadUrl(source) : viewer.loadFile(source));
}

/**
 * The vanilla building blocks: the shared building-blocks session (`ooxml-ui/pptx`) over the
 * vanilla engine, with `mount...` helpers in place of components. Mount the canvas (the deck loads
 * there), then any of the optional pieces; each returns its disposer, and `destroy()` disposes
 * whatever is still mounted. The blocks mount no dialogs or pickers: toolbar controls and canvas
 * gestures that need one raise a request (`onRequest`, `actions.pending`) the host finishes with
 * `actions`.
 */
export function createViewerBuildingBlocks(
	initial: ViewerBuildingBlocksOptions = {},
): ViewerBuildingBlocks {
	let options = initial;
	let viewer: BuildingBlocksViewer | null = null;
	let loaded: PptxViewerSource | null | undefined;
	let stopViewer: (() => void) | undefined;
	const listeners = new Set<() => void>();
	const disposers = new Set<() => void>();
	const notify = () => {
		for (const listener of [...listeners]) {
			listener();
		}
	};
	const findBar = createFindBarController({
		slides: () => viewer?.store.get().slides ?? [],
		commit: (slides) => viewer?.editor.commitElementUpdates([...slides]),
		navigate: ({ slideIndex, elementId }) => {
			if (!viewer) {
				return;
			}
			const state = viewer.store.get();
			// The editor drops the selection on a slide change, so go first, then select.
			if (state.currentSlide !== slideIndex) {
				viewer.store.set({ currentSlide: clampIndex(slideIndex, state.slides.length) });
			}
			viewer.editor.selectElements(elementId ? [elementId] : []);
		},
	});
	const session = createPptxBuildingBlockSession(
		createVanillaPort({ viewer: () => viewer, options: () => options, findBar }),
	);
	const stopRequests = session.requests.subscribe(notify);
	const stopFind = findBar.subscribe(notify);
	const subscribe = (listener: () => void) => {
		listeners.add(listener);
		return () => {
			listeners.delete(listener);
		};
	};
	const track = (dispose: () => void) => {
		const once = () => {
			if (disposers.delete(once)) {
				dispose();
			}
		};
		disposers.add(once);
		return once;
	};

	const getState = (): ViewerBuildingBlocksState => {
		if (!viewer) {
			return {
				mode: 'preview',
				loading: Boolean(options.content),
				error: null,
				slideIndex: -1,
				slideCount: 0,
				canUndo: false,
				canRedo: false,
				dirty: false,
			};
		}
		const state = viewer.store.get();
		let mode: ViewerMode = state.editable ? 'edit' : 'preview';
		if (state.presenting) {
			mode = 'present';
		} else if (state.masterViewTarget !== null) {
			mode = 'master';
		}
		return {
			mode,
			loading: state.loading,
			error: state.error,
			slideIndex: state.slides.length ? state.currentSlide : -1,
			slideCount: state.slides.length,
			canUndo: viewer.canUndo(),
			canRedo: viewer.canRedo(),
			dirty: state.dirty,
		};
	};

	const syncContent = () => {
		if (!viewer || options.content === loaded) {
			return;
		}
		loaded = options.content;
		if (loaded) {
			void load(viewer, loaded).catch(() => {
				/* The viewer reports load errors through its state and `onError`. */
			});
		}
	};

	const translate = (key: string) =>
		(viewer?.t ?? createTranslator(options.locale ?? 'en', options.messages))(key);

	return {
		get actions() {
			return session.actions();
		},
		get state() {
			return getState();
		},
		getState,
		get viewer() {
			return viewer;
		},
		subscribe,
		update(next) {
			const permission = (next.canEdit === true) !== (options.canEdit === true);
			options = next;
			if (viewer && permission) {
				viewer.setEditable(next.canEdit === true);
			}
			syncContent();
			notify();
		},
		mountCanvas(host) {
			stopViewer?.();
			viewer?.destroy();
			host.dataset.pptxBlock = 'canvas';
			loaded = options.content;
			// Permission is part of the first options, so the canvas is never briefly editable.
			viewer = new BuildingBlocksViewer(host, viewerOptions(options), (request) => {
				if (request.kind === 'find') {
					findBar.toggle();
				} else {
					session.request(request);
				}
				return true;
			});
			stopViewer = viewer.store.subscribe(notify);
			notify();
			return track(() => {
				stopViewer?.();
				stopViewer = undefined;
				viewer?.destroy();
				viewer = null;
				loaded = undefined;
				notify();
			});
		},
		mountToolbar(host, include) {
			host.dataset.pptxBlock = 'toolbar';
			host.classList.add('pptxv-bb-toolbar');
			host.setAttribute('role', 'toolbar');
			const view = createToolbarView(host.ownerDocument, (id) =>
				runBuildingBlockToolbarControl(id, session, {
					undo: () => viewer?.undo(),
					redo: () => viewer?.redo(),
					toggleFind: () => findBar.toggle(),
					zoomOut: () => viewer?.zoomOut(),
					zoomIn: () => viewer?.zoomIn(),
					zoomToFit: () => viewer?.zoomToFit(),
					previousSlide: () => viewer?.prev(),
					nextSlide: () => viewer?.next(),
				}),
			);
			const draw = () => {
				const state = getState();
				const groups = describeBuildingBlockToolbar({
					availability: session.actions().availability,
					canUndo: state.canUndo,
					canRedo: state.canRedo,
					findOpen: findBar.getSnapshot().open,
					slideIndex: state.slideIndex,
					slideCount: state.slideCount,
					hiddenActions: options.hiddenActions,
				});
				view(host, groups, translate, include);
			};
			const stop = subscribe(draw);
			draw();
			return track(() => {
				stop();
				host.replaceChildren();
			});
		},
		mountSlidesPane(host) {
			host.dataset.pptxBlock = 'slides-pane';
			return track(mountSlidesPaneBlock(host, () => ({ viewer, options }), subscribe));
		},
		destroy() {
			for (const dispose of [...disposers]) {
				dispose();
			}
			stopRequests();
			stopFind();
			listeners.clear();
		},
	};
}
