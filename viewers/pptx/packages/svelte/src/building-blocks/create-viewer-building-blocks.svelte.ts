import {
	createPptxBuildingBlockSession,
	describeBuildingBlockToolbar,
	findTextHighlightsByElement,
	registerBuildingBlockGroups,
	resolveBuildingBlockMode,
	runBuildingBlockToolbarControl,
} from 'ooxml-ui/pptx';
import type { BuildingBlockToolbarControlId } from 'ooxml-ui/pptx';
import type { PptxLayoutOption } from 'pptx-viewer-core';
import { onDestroy, untrack } from 'svelte';

import { createTranslator } from '../i18n/translator';
import { createSlidesPaneProps } from '../viewer/components/slides-pane-props';
import { controllerState } from '../viewer/state/controller-state';
import { createViewerState } from '../viewer/state/create-viewer-state.svelte';
import { toViewerStateOptions } from '../viewer/state/viewer-state-options';
import type { PowerPointViewerProps } from '../viewer/types';
import { createSvelteGroupPorts } from './group-ports';
import { routeDialogFlag } from './route-dialog-flag.svelte';
import { canEditOf, createSveltePort } from './session-port';
import type {
	ViewerBuildingBlocks,
	ViewerBuildingBlocksOptions,
	ViewerBuildingBlocksState,
} from './types';

/**
 * The viewer props the blocks hand to the viewer state (the blocks mount no chrome), as getters
 * over the live options. `content` and `canEdit` are `$derived`, so the load effect only re-runs
 * when the deck itself changes, not on every option update.
 */
function viewerProps(getOptions: () => ViewerBuildingBlocksOptions): PowerPointViewerProps {
	const content = $derived(getOptions().content);
	const canEdit = $derived(getOptions().canEdit === true);
	const option = <K extends keyof ViewerBuildingBlocksOptions>(key: K) => getOptions()[key];
	return {
		get source() {
			return content;
		},
		get editable() {
			return canEdit;
		},
		// Recovery snapshots are the host's to opt into; the blocks never write them by default.
		get autosave() {
			return option('autosave') === true;
		},
		get fileName() {
			return option('fileName');
		},
		get filePath() {
			return option('filePath');
		},
		get initialSlide() {
			return option('initialSlide');
		},
		get googleFonts() {
			return option('googleFonts');
		},
		get listConversionScope() {
			return option('listConversionScope');
		},
		get allowExternalContent() {
			return option('allowExternalContent');
		},
		get limits() {
			return option('limits');
		},
		get collaboration() {
			return option('collaboration');
		},
		get fitPadding() {
			return option('fitPadding');
		},
		get maxFitScale() {
			return option('maxFitScale');
		},
		get authorName() {
			return option('authorName') ?? option('userName');
		},
		get onload() {
			return option('onLoad');
		},
		get onerror() {
			return option('onError');
		},
		get onslidechange() {
			return option('onSlideChange');
		},
		get ondirtychange() {
			return option('onDirtyChange');
		},
		get oncontentchange() {
			return option('onContentChange');
		},
		get onselectionchange() {
			return option('onSelectionChange');
		},
	} as PowerPointViewerProps;
}

const bytesOf = (source: ViewerBuildingBlocksOptions['content']) =>
	source instanceof ArrayBuffer || source instanceof Uint8Array ? source : null;

/**
 * The Svelte building blocks: the shared building-blocks session (`ooxml-ui/pptx`) over the Svelte
 * editor's viewer state, with `BuildingBlocksCanvas`, `BuildingBlocksToolbar` and
 * `BuildingBlocksSlidesPane` as the blocks. Call it once, synchronously in the script of the
 * component that renders them (it creates the viewer state, which registers effects and context),
 * with a function returning the current options so changes apply live. The blocks mount no
 * dialogs or pickers: toolbar controls and canvas gestures that need one raise a request
 * (`onRequest`, `actions.pending`) the host finishes with `actions`.
 */
export function createViewerBuildingBlocks(
	getOptions: () => ViewerBuildingBlocksOptions,
): ViewerBuildingBlocks {
	let root = $state.raw<HTMLDivElement | undefined>();
	let stageHolder = $state.raw<HTMLDivElement | undefined>();
	let viewportWidth = $state(0);
	let viewportHeight = $state(0);
	let layouts = $state.raw<PptxLayoutOption[]>([]);
	let openPasswordSet = $state(false);
	const t = createTranslator(() => getOptions().locale ?? 'en');

	const props = viewerProps(getOptions);
	const viewerState = createViewerState(
		toViewerStateOptions(() => props, {
			t,
			getStageHolderEl: () => stageHolder,
			getRootEl: () => root,
			getViewportWidth: () => viewportWidth,
			getViewportHeight: () => viewportHeight,
			getMasterScale: () => 1,
		}),
	);
	const { editor } = viewerState;

	const session = createPptxBuildingBlockSession(
		createSveltePort({ state: viewerState, options: getOptions, layouts: () => layouts }),
	);
	registerBuildingBlockGroups(
		session,
		createSvelteGroupPorts({
			state: viewerState,
			content: () => bytesOf(getOptions().content),
			openPasswordSet: () => openPasswordSet,
			setOpenPasswordSet: (set) => {
				openPasswordSet = set;
			},
		}),
	);
	const pending = controllerState(session.requests);

	// The deck's layouts (Home > Layout, Reset): loaded per deck, for the availability input.
	$effect(() => {
		const handler = viewerState.loader.handler;
		let stale = false;
		if (handler) {
			void untrack(() => editor.slidesOps.availableLayouts())
				.then((loaded) => {
					if (!stale) {
						layouts = loaded;
					}
					return undefined;
				})
				.catch(() => undefined);
		} else {
			layouts = [];
		}
		return () => {
			stale = true;
		};
	});
	$effect(() => {
		openPasswordSet = viewerState.loader.isPasswordProtected;
	});

	// The canvas opens its dialogs by flag (Ctrl+K, double-clicking an equation, `?`): nothing
	// renders them here, so each flag becomes a host request and is cleared.
	routeDialogFlag(
		() => viewerState.controller.hyperlinkOpen,
		() => {
			viewerState.controller.hyperlinkOpen = false;
			session.request({ kind: 'hyperlink' });
		},
	);
	routeDialogFlag(
		() => Boolean(editor.equationOps.editingId),
		() => {
			const existingOmml = editor.equationOps.omml;
			editor.equationOps.close();
			session.request({ kind: 'equation', existingOmml });
		},
	);
	// The shortcut help is a toggle the host mirrors: every change of its flag is an event.
	let shortcutsShown = false;
	$effect(() => {
		const open = viewerState.parityUi.shortcutsOpen;
		if (open !== shortcutsShown) {
			shortcutsShown = open;
			untrack(() => session.announce({ kind: 'shortcuts', open }, open));
		}
	});

	let destroyed = false;
	const destroy = (): void => {
		if (!destroyed) {
			destroyed = true;
			viewerState.destroy();
		}
	};
	onDestroy(destroy);

	const getState = (): ViewerBuildingBlocksState => ({
		mode: resolveBuildingBlockMode(viewerState.viewerMode, canEditOf(viewerState)),
		loading: viewerState.loader.loading,
		error: viewerState.loader.error,
		slideIndex: editor.slides.length ? viewerState.viewer.current : -1,
		slideCount: editor.slides.length,
		canUndo: editor.canUndo,
		canRedo: editor.canRedo,
		dirty: editor.dirty,
	});

	return {
		get actions() {
			// The pending request is read here so a template that reads `actions` tracks it.
			void pending.current;
			return session.actions();
		},
		get state() {
			return getState();
		},
		get toolbar() {
			const state = getState();
			return describeBuildingBlockToolbar({
				availability: session.actions().availability,
				canUndo: state.canUndo,
				canRedo: state.canRedo,
				findOpen: viewerState.findReplace.open,
				slideIndex: state.slideIndex,
				slideCount: state.slideCount,
				hiddenActions: getOptions().hiddenActions,
			});
		},
		runToolbarControl(id: BuildingBlockToolbarControlId) {
			runBuildingBlockToolbarControl(id, session, {
				undo: () => editor.undo(),
				redo: () => editor.redo(),
				toggleFind: () => viewerState.findReplace.toggle(),
				zoomOut: () => viewerState.deck.zoomOut(),
				zoomIn: () => viewerState.deck.zoomIn(),
				zoomToFit: () => {
					viewerState.viewer.zoomPercent = null;
				},
				previousSlide: () => viewerState.deck.goPrev(),
				nextSlide: () => viewerState.deck.goNext(),
			});
		},
		slidesPaneProps: createSlidesPaneProps(viewerState, {
			onopenlayoutgallery: (request) => getOptions().onOpenLayoutGallery?.(request),
		}),
		get findHighlights() {
			return findTextHighlightsByElement(
				viewerState.findReplace.results,
				viewerState.findReplace.index,
				viewerState.viewer.current,
			);
		},
		t: (key) => t(key),
		viewerState,
		session,
		canvas: {
			setRoot(element) {
				root = element;
			},
			setStageHolder(element) {
				stageHolder = element;
			},
			setViewportSize(size) {
				viewportWidth = size.width;
				viewportHeight = size.height;
			},
		},
		destroy,
	};
}
