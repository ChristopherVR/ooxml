import {
	createBuildingBlockShell,
	createPptxBuildingBlockSession,
	describeBuildingBlockToolbar,
	findTextHighlightsByElement,
	registerBuildingBlockGroups,
	resolveBuildingBlockMode,
	resolveCollaborationConfig,
	runBuildingBlockToolbarControl,
} from 'ooxml-ui/pptx';
import type { BuildingBlockToolbarControlId } from 'ooxml-ui/pptx';
import type { PptxLayoutOption } from 'pptx-viewer-core';
import { onDestroy, untrack } from 'svelte';

import { createTranslator } from '../i18n/translator';
import { createSlidesPaneProps } from '../viewer/components/slides-pane-props';
import { controllerState } from '../viewer/state/controller-state';
import { createViewerState } from '../viewer/state/create-viewer-state.svelte';
import { provideElementRenderErrorReporter } from '../viewer/state/element-render-errors';
import {
	useCustomizationConstraints,
	useViewerCustomizationRoot,
} from '../viewer/state/viewer-customization.svelte';
import { toViewerStateOptions } from '../viewer/state/viewer-state-options';
import { createSvelteGroupPorts } from './group-ports';
import { routeDialogFlag } from './route-dialog-flag.svelte';
import { canEditOf, createSveltePort } from './session-port';
import type {
	ViewerBuildingBlocks,
	ViewerBuildingBlocksOptions,
	ViewerBuildingBlocksState,
} from './types';
import { viewerProps } from './viewer-props.svelte';

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
	// The host `allowExternalContent` and `googleFonts` lock their Trust Center settings, live.
	const customization = useViewerCustomizationRoot(
		() => undefined,
		() => getOptions().allowExternalContent,
		() => getOptions().googleFonts,
	);
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
	useCustomizationConstraints(customization, viewerState.optionsState);
	// Context of the creating component, so the canvas and slides pane blocks below report here.
	provideElementRenderErrorReporter(() => getOptions().onElementRenderError);

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

	// The shell over the viewer state's collaboration session; the same object until it changes.
	const shell = createBuildingBlockShell(() => {
		const { collab, loader } = viewerState;
		return {
			shell: {
				authorizedCanEdit: getOptions().canEdit === true,
				configured:
					Boolean(collab.activeCollaboration) ||
					Boolean(resolveCollaborationConfig(getOptions().collaboration)),
				readOnly: collab.readOnly,
				sourcePending: Boolean(getOptions().content) && loader.loading,
				sourceError: Boolean(loader.error),
				status: collab.status,
				remoteUsers: collab.remotePresences,
			},
			collaboration: collab.state,
		};
	});

	const getState = (): ViewerBuildingBlocksState => ({
		...shell.current(),
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
