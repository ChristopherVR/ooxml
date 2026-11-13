import type {
	BuildingBlockToolbarGroupId,
	CollaborationShellState,
	CollaborationState,
	UseViewerBuildingBlocksActionsOptions,
	ViewerBuildingBlockActions,
	ViewerMode,
} from 'ooxml-ui/pptx';

import type { PptxViewerSource } from '../viewer/load';
import type { PptxViewer } from '../viewer/PptxViewer';
import type { PptxViewerOptions } from '../viewer/types';
import type { SlidesPaneLayoutGalleryRequest } from '../viewer/ui/slides-pane';

/** Viewer options the building blocks decide themselves (they mount no chrome). */
type ChromeOptions =
	| 'source'
	| 'editable'
	| 'readOnly'
	| 'showToolbar'
	| 'showThumbnails'
	| 'showFormatToolbar'
	| 'showInspector'
	| 'showCompatibilityToasts';

export interface ViewerBuildingBlocksOptions
	extends Omit<PptxViewerOptions, ChromeOptions>, UseViewerBuildingBlocksActionsOptions {
	/** The presentation (bytes, a Blob or a URL), or null while no file is open. */
	content?: PptxViewerSource | null | undefined;
	/** Whether editing is enabled. Defaults to false (view-only), as React's `canEdit`. */
	canEdit?: boolean | undefined;
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
	/** Host-owned workflows (pickers, dialogs, Find / Replace, layouts); see the vanilla docs. */
	readonly actions: ViewerBuildingBlockActions;
	/** The current state; same value `getState()` returns. */
	readonly state: ViewerBuildingBlocksState;
	getState(): ViewerBuildingBlocksState;
	/** The full viewer instance behind the canvas, for its imperative API; null until mounted. */
	readonly viewer: PptxViewer | null;
	/** Call `listener` after every change of state, actions or options; returns the unsubscribe. */
	subscribe(listener: () => void): () => void;
	/**
	 * Apply new options to the mounted canvas, keeping the open deck: `content`, `canEdit`,
	 * callbacks, `locale`, `theme`, `customization`, `googleFonts`, `allowExternalContent` and
	 * `collaboration` (join, rejoin or leave) at once; `limits` and `initialSlide` on the next load.
	 */
	update(next: ViewerBuildingBlocksOptions): void;
	/** Mount the slide canvas (the deck loads here). Returns its disposer. */
	mountCanvas(host: HTMLElement): () => void;
	/** Mount the compact toolbar; `include` picks groups. Returns its disposer. */
	mountToolbar(host: HTMLElement, include?: readonly BuildingBlockToolbarGroupId[]): () => void;
	/** Mount the slide thumbnail rail. Returns its disposer. */
	mountSlidesPane(host: HTMLElement): () => void;
	/** Dispose every mounted block and the viewer. */
	destroy(): void;
}
