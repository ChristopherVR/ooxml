import { resolveBuildingBlockMode } from 'ooxml-ui/pptx';
import type { ViewerMode } from 'ooxml-ui/pptx';
import { onBeforeUnmount, onMounted, watch } from 'vue';

import type { BuildingBlocksBridge, BuildingBlocksEngine } from './bridge';
import type { BuildingBlocksEngineContext as Context } from './engine-context';
import { createEngineGroupPorts, createEngineSessionPort } from './engine-ports';

/**
 * The dialogs the viewer would open from its own flags (hyperlink, equation, SmartArt, the file and
 * slide show dialogs): the building blocks mount none, so each flag is cleared at once and raised
 * as a host request instead. Sync watchers, so the dialog never renders.
 */
/**
 * Clear a dialog flag and raise its host request when it turns on. A sync watcher on the flag
 * alone: raising notifies the session's store, and a watcher that also tracked anything that
 * notification touches could run again while the flag is still on (the loop React 18 hit). The
 * flag is cleared first, so the request is raised once per opening.
 */
export function routeDialogFlag(flag: () => boolean, clear: () => void, raise: () => void) {
	return watch(
		flag,
		(open) => {
			if (open) {
				clear();
				raise();
			}
		},
		{ flush: 'sync', immediate: true },
	);
}

function routeDialogFlags(c: Context, bridge: BuildingBlocksBridge): void {
	const route = routeDialogFlag;
	const { request } = bridge.host;
	route(
		() => c.hyperlink.hyperlinkOpen.value,
		() => {
			c.hyperlink.hyperlinkOpen.value = false;
		},
		() => request({ kind: 'hyperlink' }),
	);
	// The equation being edited is read before the editor is closed (closing clears it).
	let existingOmml: Record<string, unknown> | null = null;
	route(
		() => c.insertDialogs.showEquationEditor.value,
		() => {
			existingOmml = c.insertDialogs.editingEquationOmml.value ?? null;
			c.insertDialogs.closeEquationEditor();
		},
		() => request({ kind: 'equation', existingOmml }),
	);
	route(
		() => c.insertDialogs.showInsertSmartArt.value,
		() => {
			c.insertDialogs.showInsertSmartArt.value = false;
		},
		() => request({ kind: 'smartArt' }),
	);
	route(
		() => c.password.showPasswordDialog.value,
		() => {
			c.password.showPasswordDialog.value = false;
		},
		() => request({ kind: 'passwordProtection' }),
	);
	route(
		() => c.documentProperties.propertiesOpen.value,
		() => {
			c.documentProperties.propertiesOpen.value = false;
		},
		() => request({ kind: 'documentProperties' }),
	);
	route(
		() => c.fontEmbedding.showFontEmbedding.value,
		() => {
			c.fontEmbedding.showFontEmbedding.value = false;
		},
		() => request({ kind: 'fontEmbedding' }),
	);
	route(
		() => c.signatureWorkflow.showSignatures.value,
		() => {
			c.signatureWorkflow.showSignatures.value = false;
		},
		() => request({ kind: 'digitalSignatures' }),
	);
	route(
		() => c.slideShow.showSetUpSlideShow.value,
		() => {
			c.slideShow.showSetUpSlideShow.value = false;
		},
		() => request({ kind: 'setUpSlideShow' }),
	);
	route(
		() => c.customShowsWiring.showCustomShows.value,
		() => {
			c.customShowsWiring.showCustomShows.value = false;
		},
		() => request({ kind: 'customShows' }),
	);
	// Broadcast is unsupported (`actions.broadcast`): the flag is dropped, no request is raised.
	route(
		() => c.broadcastOpen.value,
		() => {
			c.broadcastOpen.value = false;
		},
		() => undefined,
	);
	// The shortcut help is a toggle the host mirrors: every change of its flag is an event.
	watch(c.showShortcuts, (open) => bridge.host.announceShortcuts(open), { flush: 'sync' });
}

function createEngine(c: Context): BuildingBlocksEngine {
	return {
		session: createEngineSessionPort(c),
		groups: createEngineGroupPorts(c),
		state: () => ({
			mode: resolveBuildingBlockMode(c.mode.value as ViewerMode, c.canEditBase()),
			canEdit: c.canEdit.value,
			loading: c.loading.value,
			error: c.error.value,
			slideIndex: c.slideCount.value ? c.activeSlideIndex.value : -1,
			slideCount: c.slideCount.value,
			canUndo: c.history.canUndo.value,
			canRedo: c.history.canRedo.value,
			dirty: c.isDirty.value,
		}),
		view: {
			undo: c.history.undo,
			redo: c.history.redo,
			toggleFind: () => {
				c.findOpen.value = !c.findOpen.value;
			},
			zoomOut: c.zoomOut,
			zoomIn: c.zoomIn,
			zoomToFit: c.zoomReset,
			previousSlide: c.goPrev,
			nextSlide: c.goNext,
		},
		closeShortcuts: () => {
			c.showShortcuts.value = false;
		},
		slidesPane: () => ({
			slides: c.mergedSlides.value,
			activeSlideIndex: c.activeSlideIndex.value,
			canvasSize: c.canvasSize.value,
			mediaDataUrls: c.mediaDataUrls.value,
			canEdit: c.canEdit.value,
			sectionGroups: c.hasSections.value ? c.sectionOps.slidesBySection.value : [],
			headerFooter: c.headerFooter.value,
			customProperties: c.customProperties.value,
			firstSlideNumber: c.firstSlideNumber.value,
			theme: c.theme.value,
			tableStyleMap: c.tableStyleMap.value,
			goTo: (index) => {
				c.activeSlideIndex.value = index;
			},
			toggleSlideHidden: c.toggleSlideHidden,
			slideOps: c.slideOps,
			sectionOps: c.sectionOps,
		}),
		findHighlights: () => ({
			matches: c.find.matches.value,
			current: c.find.currentMatch.value,
			slideIndex: c.activeSlideIndex.value,
		}),
	};
}

/**
 * Called by `PowerPointViewer` when a building-blocks canvas provides a bridge: routes the
 * viewer's own dialogs to host requests and publishes the engine the blocks read.
 */
export function useBuildingBlocksEngine(bridge: BuildingBlocksBridge, c: Context): void {
	routeDialogFlags(c, bridge);
	const engine = createEngine(c);
	let detach: (() => void) | undefined;
	onMounted(() => {
		detach = bridge.attach(engine);
	});
	onBeforeUnmount(() => detach?.());
}
