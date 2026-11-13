import {
	createBuildingBlockShell,
	createPptxBuildingBlockSession,
	registerBuildingBlockGroups,
	resolveCollaborationConfig,
} from 'ooxml-ui/pptx';
import type {
	BuildingBlockAvailabilityInput,
	BuildingBlockGroupPorts,
	BuildingBlockSessionPort,
	ViewerMode,
} from 'ooxml-ui/pptx';
import { computed, onScopeDispose, toValue } from 'vue';
import type { MaybeRefOrGetter } from 'vue';

import { useController } from '../viewer/composables/useController';
import { createBuildingBlocksBridge } from './bridge';
import type { BuildingBlocksEngine } from './bridge';
import type {
	ViewerBuildingBlocks,
	ViewerBuildingBlocksOptions,
	ViewerBuildingBlocksState,
} from './types';

const NOT_MOUNTED: BuildingBlockAvailabilityInput = {
	ready: false,
	hasActiveSlide: false,
	canEdit: false,
	editing: false,
	canInsert: false,
	hasSelection: false,
	selectionHasHyperlink: false,
	selectionHasEquation: false,
	layoutCount: 0,
	hasCurrentLayout: false,
	hasHandler: false,
};

/** Group ports that forward to whatever engine is mounted (every port is a bag of functions). */
function lazyGroupPorts(read: () => BuildingBlockGroupPorts | undefined): BuildingBlockGroupPorts {
	const idle = {
		input: () => ({
			openPasswordSet: false,
			recommendation: { kind: null },
			hasModifyVerifier: false,
			customProperties: [],
			embedFonts: false,
			embeddedFamilies: [],
			usedFamilies: [],
			descriptor: { canEmbed: false },
		}),
		customProperties: () => [],
		current: () => ({ core: {}, custom: [], app: {} }),
		signed: () => false,
		count: () => 0,
		content: () => null,
		presentationProperties: () => ({}),
		slideCount: () => 0,
		customShows: () => [],
		slides: () => [],
		activeId: () => null,
		open: () => false,
		translate: (key: string) => key,
	} as Record<string, (...args: never[]) => unknown>;
	const group = <K extends keyof BuildingBlockGroupPorts>(key: K): BuildingBlockGroupPorts[K] =>
		new Proxy({} as BuildingBlockGroupPorts[K], {
			get:
				(_target, member: string) =>
				(...args: unknown[]) => {
					const port = read()?.[key] as unknown as Record<string, (...a: unknown[]) => unknown>;
					return port ? port[member]?.(...args) : idle[member]?.();
				},
		});
	return {
		passwordProtection: group('passwordProtection'),
		documentProperties: group('documentProperties'),
		fontEmbedding: group('fontEmbedding'),
		digitalSignatures: group('digitalSignatures'),
		setUpSlideShow: group('setUpSlideShow'),
		customShows: group('customShows'),
		shortcuts: group('shortcuts'),
	};
}

/**
 * The Vue building blocks: the shared building-blocks session (`ooxml-ui/pptx`) over the Vue
 * viewer. Render `<BuildingBlocksCanvas :blocks>` (the deck loads there), then any of
 * `<BuildingBlocksToolbar>` and `<BuildingBlocksSlidesPane>`. The blocks mount no dialogs or
 * pickers: controls and gestures that need one raise a request (`onRequest`, `actions.pending`)
 * the host finishes with `actions`.
 */
export function useViewerBuildingBlocks(
	input: MaybeRefOrGetter<ViewerBuildingBlocksOptions>,
): ViewerBuildingBlocks {
	const options = computed(() => toValue(input));
	const bridge = createBuildingBlocksBridge({
		request: (request) => void session.request(request),
		announceShortcuts: (open) => session.announce({ kind: 'shortcuts', open }, open),
	});
	const engine = (): BuildingBlocksEngine | null => bridge.engine.value;
	const port: BuildingBlockSessionPort = {
		availabilityInput: () => engine()?.session.availabilityInput() ?? NOT_MOUNTED,
		fileLimits: () => options.value.fileLimits,
		onRequest: (request) => options.value.onRequest?.(request),
		onCancel: () => engine()?.closeShortcuts(),
		target: () => engine()?.session.target() ?? { deck: null, slideId: undefined, slideIndex: -1 },
		canvasSize: () => engine()?.session.canvasSize() ?? { width: 0, height: 0 },
		selectedElement: () => engine()?.session.selectedElement() ?? null,
		slideCount: () => engine()?.session.slideCount() ?? 0,
		activeSlide: () => engine()?.session.activeSlide(),
		scopedLayouts: () => engine()?.session.scopedLayouts() ?? [],
		addElement: (element) => engine()?.session.addElement(element),
		insertMediaFile: async (file, sameTarget) => {
			const current = engine();
			return current
				? current.session.insertMediaFile(file, sameTarget)
				: { ok: false, code: 'not-ready' };
		},
		applyHyperlink: (data) =>
			engine()?.session.applyHyperlink(data) ?? { ok: false, code: 'not-ready' },
		removeHyperlink: () => engine()?.session.removeHyperlink() ?? { ok: false, code: 'not-ready' },
		insertSmartArt: (layout, items) => engine()?.session.insertSmartArt(layout, items),
		insertChart: (kind, chartOptions) => engine()?.session.insertChart(kind, chartOptions),
		insertEquation: (omml) => engine()?.session.insertEquation(omml),
		updateEquation: (omml) => engine()?.session.updateEquation(omml),
		findReplace: () => {
			const current = engine();
			if (current) {
				return current.session.findReplace();
			}
			return {
				open: false,
				setOpen: () => undefined,
				query: '',
				replacement: '',
				options: {
					matchCase: false,
					includeNotes: false,
					includeChartTitles: false,
					includeChartData: false,
				},
				matches: [],
				currentIndex: -1,
				setQuery: () => undefined,
				setReplacement: () => undefined,
				setOptions: () => undefined,
				find: () => undefined,
				navigate: () => undefined,
				replaceCurrent: () => undefined,
				replaceAll: () => undefined,
			};
		},
		applyLayout: async (path) => {
			const current = engine();
			if (!current) {
				throw new Error(`Layout ${path} was not applied`);
			}
			await current.session.applyLayout(path);
		},
	};
	const session = createPptxBuildingBlockSession(port);
	const stopGroups = registerBuildingBlockGroups(
		session,
		lazyGroupPorts(() => engine()?.groups),
	);
	const requests = useController(session.requests);
	onScopeDispose(stopGroups);

	const actions = computed(() => {
		// Track the pending request and the mounted engine; `session.actions()` reads the rest.
		void requests.value;
		void bridge.engine.value;
		return session.actions();
	});
	const shell = createBuildingBlockShell(
		() =>
			engine()?.shell({
				canEdit: options.value.canEdit === true,
				configured: Boolean(resolveCollaborationConfig(options.value.collaboration)),
			}) ?? null,
	);
	const shellView = computed(() => shell.current());
	const state = computed<ViewerBuildingBlocksState>(() => {
		const current = engine()?.state();
		const { shellState, collaborationState } = shellView.value;
		if (!current) {
			return {
				shellState,
				collaborationState,
				mode: 'preview',
				loading: Boolean(options.value.content),
				error: null,
				slideIndex: -1,
				slideCount: 0,
				canUndo: false,
				canRedo: false,
				dirty: false,
			};
		}
		const { canEdit: _canEdit, ...rest } = current;
		return { ...rest, mode: current.mode as ViewerMode, shellState, collaborationState };
	});
	return {
		actions,
		state,
		mode: computed(() => state.value.mode),
		loading: computed(() => state.value.loading),
		error: computed(() => state.value.error),
		shellState: computed(() => shellView.value.shellState),
		collaborationState: computed(() => shellView.value.collaborationState),
		viewer: bridge.viewer,
		options,
		session,
		bridge,
	};
}
