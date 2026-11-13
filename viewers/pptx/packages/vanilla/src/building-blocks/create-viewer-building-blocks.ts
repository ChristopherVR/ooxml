import {
	createControllerStore,
	createFindBarController,
	createPptxBuildingBlockSession,
	describeBuildingBlockToolbar,
	registerBuildingBlockGroups,
	runBuildingBlockToolbarControl,
} from 'ooxml-ui/pptx';
import type { FindResult } from 'ooxml-ui/pptx';

import { createTranslator } from '../viewer/i18n';
import type { PptxViewerSource } from '../viewer/load';
import { clampSlideIndex as clampIndex } from '../viewer/state';
import { viewerModeOf } from '../viewer/state-sync';
import { BuildingBlocksViewer } from './blocks-viewer';
import { createVanillaGroupPorts } from './group-ports';
import { createVanillaPort } from './session-port';
import { mountSlidesPaneBlock } from './slides-pane-block';
import { createToolbarView } from './toolbar-view';
import type {
	ViewerBuildingBlocks,
	ViewerBuildingBlocksOptions,
	ViewerBuildingBlocksState,
} from './types';
import { applyOptionChanges, liveViewerOptions, loadSource } from './viewer-options';

const bytesOf = (source: PptxViewerSource | null | undefined) =>
	source instanceof ArrayBuffer || source instanceof Uint8Array ? source : null;

/**
 * The vanilla building blocks: the shared building-blocks session (`ooxml-ui/pptx`) over the
 * vanilla engine, with `mount...` helpers in place of components. Mount the canvas (the deck loads
 * there), then any of the optional pieces; each returns its disposer, and `destroy()` disposes
 * whatever is still mounted. The blocks mount no dialogs or pickers: toolbar controls and canvas
 * gestures that need one raise a request (`onRequest`, `actions.pending`) the host finishes with
 * `actions`.
 */
/** The matches a closed host find bar marks: none (one array, so the store sees no change). */
const NO_MATCHES: readonly FindResult[] = [];

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
	// The shortcut help is the host's: the blocks keep only whether it is open.
	const shortcuts = createControllerStore(false);
	const translate = (key: string) =>
		(viewer?.t ?? createTranslator(options.locale ?? 'en', options.messages))(key);
	const session = createPptxBuildingBlockSession(
		createVanillaPort({
			viewer: () => viewer,
			options: () => options,
			findBar,
			onCancel: () => setShortcutsOpen(false),
		}),
	);
	const setShortcutsOpen = (open: boolean) => {
		if (shortcuts.getSnapshot() !== open) {
			shortcuts.set(open);
			session.announce({ kind: 'shortcuts', open }, open);
		}
	};
	registerBuildingBlockGroups(
		session,
		createVanillaGroupPorts({
			viewer: () => viewer,
			content: () => bytesOf(loaded),
			shortcutsOpen: shortcuts.getSnapshot,
			setShortcutsOpen,
			translate,
		}),
	);
	const stops = [
		session.requests.subscribe(notify),
		findBar.subscribe(notify),
		shortcuts.subscribe(notify),
	];
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
		return {
			// Already the effective mode (`preview` while not editable), as `resolveBuildingBlockMode`.
			mode: viewerModeOf(state),
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
			void loadSource(viewer, loaded).catch(() => {
				/* The viewer reports load errors through its state and `onError`. */
			});
		}
	};

	const mountCanvas = (host: HTMLElement) => {
		stopViewer?.();
		viewer?.destroy();
		host.dataset.pptxBlock = 'canvas';
		loaded = options.content;
		// Permission is part of the first options, so the canvas is never briefly editable.
		const mounted = new BuildingBlocksViewer(
			host,
			liveViewerOptions(() => options),
			(request) => {
				if (request.kind === 'find') {
					findBar.toggle();
				} else if (request.kind === 'shortcuts') {
					setShortcutsOpen(!shortcuts.getSnapshot());
				} else {
					session.request(request);
				}
				return true;
			},
		);
		viewer = mounted;
		// The host find bar's matches: the viewer's text renderer marks them on the canvas.
		const highlight = () => {
			const { open, session: find } = findBar.getSnapshot();
			const matches = open ? find.matches : NO_MATCHES;
			const state = mounted.store.get();
			if (state.findMatches !== matches || state.findMatchIndex !== find.current) {
				mounted.store.set({ findMatches: matches, findMatchIndex: find.current });
			}
		};
		highlight();
		const stopStore = mounted.store.subscribe(notify);
		const stopFind = findBar.subscribe(highlight);
		stopViewer = () => {
			stopStore();
			stopFind();
		};
		notify();
		return track(() => {
			stopViewer?.();
			stopViewer = undefined;
			viewer?.destroy();
			viewer = null;
			loaded = undefined;
			notify();
		});
	};

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
			const previous = options;
			options = next;
			if (viewer) {
				applyOptionChanges(viewer, previous, next);
			}
			syncContent();
			notify();
		},
		mountCanvas,
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
			for (const stop of stops) {
				stop();
			}
			listeners.clear();
		},
	};
}
