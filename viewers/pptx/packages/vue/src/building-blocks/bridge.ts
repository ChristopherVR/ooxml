import type {
	BuildingBlockShellInput,
	BuildingBlockGroupPorts,
	BuildingBlockRequest,
	BuildingBlockSessionPort,
	BuildingBlockToolbarView,
} from 'ooxml-ui/pptx';
import type {
	ParsedTableStyleMap,
	PptxCustomProperty,
	PptxHeaderFooter,
	PptxSlide,
	PptxTheme,
} from 'pptx-viewer-core';
import { shallowRef } from 'vue';
import type { InjectionKey, ShallowRef } from 'vue';

import type { UseSectionOperationsResult } from '../viewer/composables/useSectionOperations';
import type { SectionGroup } from '../viewer/composables/useSectionOperations';
import type { UseSlideOperationsResult } from '../viewer/composables/useSlideOperations';
import type { PowerPointViewerExpose } from '../viewer/types';
import type { ViewerBuildingBlocksState } from './types';

/**
 * What the viewer behind the canvas block publishes, read at call time: the session port over its
 * own editor state, the dialog group ports, the view operations and what the slides pane draws.
 * Every getter reads the viewer's refs, so reading them from a computed or a render tracks them.
 */
export interface BuildingBlocksEngine {
	/** The members of the session port the viewer owns (the host's options are added around them). */
	readonly session: Omit<BuildingBlockSessionPort, 'fileLimits' | 'onRequest' | 'onCancel'>;
	readonly groups: BuildingBlockGroupPorts;
	state(): Omit<ViewerBuildingBlocksState, 'mode' | 'shellState' | 'collaborationState'> & {
		mode: string;
		canEdit: boolean;
	};
	readonly view: BuildingBlockToolbarView;
	/** The shortcut help flag's mirror: close it (`actions.cancel()`). */
	closeShortcuts(): void;
	/** What the slides pane block draws and which operations it maps its events onto. */
	slidesPane(): SlidesPaneBinding;
	/** The shell inputs from the viewer's collaboration session (`host` is the host's permission and config). */
	shell(host: { canEdit: boolean; configured: boolean }): BuildingBlockShellInput;
}

/** The viewer's slide rail inputs and operations, as `ViewerSlideRail` reads them. */
export interface SlidesPaneBinding {
	slides: readonly PptxSlide[];
	activeSlideIndex: number;
	canvasSize: { width: number; height: number };
	mediaDataUrls: Map<string, string>;
	canEdit: boolean;
	sectionGroups: SectionGroup[];
	headerFooter: PptxHeaderFooter | undefined;
	customProperties: readonly PptxCustomProperty[] | undefined;
	firstSlideNumber: number | undefined;
	theme: PptxTheme | undefined;
	tableStyleMap: ParsedTableStyleMap | undefined;
	goTo(index: number): void;
	toggleSlideHidden(index: number): void;
	slideOps: UseSlideOperationsResult;
	sectionOps: UseSectionOperationsResult;
}

/** What the viewer calls on the host side: requests and the mirrored shortcut help. */
export interface BuildingBlocksHost {
	request(request: BuildingBlockRequest): void;
	announceShortcuts(open: boolean): void;
}

export interface BuildingBlocksBridge {
	readonly engine: Readonly<ShallowRef<BuildingBlocksEngine | null>>;
	readonly host: BuildingBlocksHost;
	/** The viewer's imperative handle, set by the canvas block. */
	readonly viewer: ShallowRef<PowerPointViewerExpose | null>;
	/** The viewer registers its engine; returns the function that unregisters it. */
	attach(engine: BuildingBlocksEngine): () => void;
}

/** Provided by the canvas block; `PowerPointViewer` injects it to publish its engine. */
export const BuildingBlocksBridgeKey: InjectionKey<BuildingBlocksBridge> = Symbol(
	'pptx-vue-building-blocks',
);

export function createBuildingBlocksBridge(host: BuildingBlocksHost): BuildingBlocksBridge {
	const engine = shallowRef<BuildingBlocksEngine | null>(null);
	return {
		engine,
		host,
		viewer: shallowRef<PowerPointViewerExpose | null>(null),
		attach(next) {
			engine.value = next;
			return () => {
				if (engine.value === next) {
					engine.value = null;
				}
			};
		},
	};
}
