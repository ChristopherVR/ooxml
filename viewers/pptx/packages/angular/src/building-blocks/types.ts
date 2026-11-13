import type {
	CollaborationShellState,
	CollaborationState,
	GoogleFontsOption,
	ListConversionScope,
	PptxLoadLimits,
	SlidesPaneLayoutGalleryRequest,
	ToolbarActionId,
	UseViewerBuildingBlocksActionsOptions,
	ViewerMode,
} from 'ooxml-ui/pptx';

import type { CollaborationInput } from '../viewer/types';

/** What a host passes to `createViewerBuildingBlocks(() => options)` or `blocks.update(options)`. */
export interface ViewerBuildingBlocksOptions extends UseViewerBuildingBlocksActionsOptions {
	/** The presentation bytes, or null while no file is open. */
	content?: Uint8Array | ArrayBuffer | null | undefined;
	/** Whether editing is enabled. Defaults to false (view-only). */
	canEdit?: boolean | undefined;
	/** Optional collaboration, including a host-owned document and presence session. */
	collaboration?: CollaborationInput | undefined;
	/** Zero-based slide shown after each load (applies to the next load). */
	initialSlide?: number | undefined;
	/** Bounds on loading a deck (applies to the next load). */
	limits?: PptxLoadLimits | undefined;
	googleFonts?: GoogleFontsOption | undefined;
	allowExternalContent?: boolean | undefined;
	listConversionScope?: ListConversionScope | undefined;
	/** Toolbar controls to hide (`undo`, `redo`, `insert`, `zoom`, `navigation` apply). */
	hiddenActions?: readonly ToolbarActionId[] | undefined;
	onDirtyChange?: ((dirty: boolean) => void) | undefined;
	/** Fired by the "Layout" entry of a slide's context menu in the slides pane. */
	onOpenLayoutGallery?: ((request: SlidesPaneLayoutGalleryRequest) => void) | undefined;
}

/** The state a custom shell renders around the blocks (a status line, Save enablement). */
export interface ViewerBuildingBlocksState {
	/** `preview` while the blocks may not edit, otherwise the viewer's mode. */
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
