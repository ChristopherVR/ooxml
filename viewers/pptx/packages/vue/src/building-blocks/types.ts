import type {
	BuildingBlockToolbarGroupId,
	CollaborationShellState,
	CollaborationState,
	UseViewerBuildingBlocksActionsOptions,
	ViewerBuildingBlockActions,
	PptxBuildingBlockSession,
	SlidesPaneLayoutGalleryRequest,
	ViewerMode,
} from 'ooxml-ui/pptx';
import type { ComputedRef, ShallowRef } from 'vue';

import type { PowerPointViewerExpose, PowerPointViewerProps } from '../viewer/types';
import type { BuildingBlocksBridge } from './bridge';

/** Viewer props the building blocks decide themselves (they mount no chrome). */
type ChromeProps =
	| 'content'
	| 'canEdit'
	| 'readOnly'
	| 'showToolbar'
	| 'showThumbnails'
	| 'showCompatibilityToasts';

export interface ViewerBuildingBlocksOptions
	extends Omit<PowerPointViewerProps, ChromeProps>, UseViewerBuildingBlocksActionsOptions {
	/** The presentation bytes, or null while no file is open. */
	content?: Uint8Array | ArrayBuffer | null | undefined;
	/** Whether editing is enabled. Defaults to false (view-only), as React's `canEdit`. */
	canEdit?: boolean | undefined;
	onContentChange?: ((content: Uint8Array) => void) | undefined;
	onDirtyChange?: ((dirty: boolean) => void) | undefined;
	onActiveSlideChange?: ((index: number) => void) | undefined;
	onModeChange?: ((mode: string) => void) | undefined;
	onZoomChange?: ((zoom: number) => void) | undefined;
	onSelectionChange?: ((ids: string[]) => void) | undefined;
	onSlideCountChange?: ((count: number) => void) | undefined;
	/**
	 * Fired by the "Layout" entry of a slide's context menu in the slides pane, after that slide is
	 * made active; the host owns rendering the Layout gallery.
	 */
	onOpenLayoutGallery?: ((request: SlidesPaneLayoutGalleryRequest) => void) | undefined;
}

/** The state a custom shell renders around the blocks (a status line, Save enablement). */
export interface ViewerBuildingBlocksState {
	/** `preview` until a canvas is mounted and the host grants editing. */
	mode: ViewerMode;
	loading: boolean;
	error: string | null;
	/** Zero-based active slide, -1 before a deck loads. */
	slideIndex: number;
	slideCount: number;
	canUndo: boolean;
	canRedo: boolean;
	dirty: boolean;
	/**
	 * Effective edit permission plus connection and presence counts, as React's `shellState`;
	 * pass it to `describeCollaborationShellState` for a localised status line.
	 */
	shellState: CollaborationShellState;
	/** The shared `CollaborationState` of the room the canvas is in (`COLLABORATION_OFF` when none). */
	collaborationState: CollaborationState;
}

export interface ViewerBuildingBlocks {
	/** Host-owned workflows (pickers, dialogs, Find / Replace, layouts); see the Vue docs. */
	readonly actions: ComputedRef<ViewerBuildingBlockActions>;
	/** The state for a custom shell. */
	readonly state: ComputedRef<ViewerBuildingBlocksState>;
	/** The effective mode (`preview` while the blocks may not edit). */
	readonly mode: ComputedRef<ViewerMode>;
	readonly loading: ComputedRef<boolean>;
	readonly error: ComputedRef<string | null>;
	/** `state.shellState`, as its own ref. */
	readonly shellState: ComputedRef<CollaborationShellState>;
	/** `state.collaborationState`, as its own ref. */
	readonly collaborationState: ComputedRef<CollaborationState>;
	/** The viewer's imperative handle (save, export, zoom, ...); null until the canvas mounts. */
	readonly viewer: Readonly<ShallowRef<PowerPointViewerExpose | null>>;
	/** The options the blocks currently run with. */
	readonly options: ComputedRef<ViewerBuildingBlocksOptions>;
	/** The shared session behind `actions` (for hosts that register further groups). */
	readonly session: PptxBuildingBlockSession;
	/** Wiring between the three blocks; used by the block components, not part of the API. */
	readonly bridge: BuildingBlocksBridge;
}

export type { BuildingBlockToolbarGroupId };
