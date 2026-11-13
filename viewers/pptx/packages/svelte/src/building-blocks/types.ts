import type {
	BuildingBlockToolbarControlId,
	BuildingBlockToolbarGroup,
	SlideFindHighlights,
	PptxBuildingBlockSession,
	ToolbarActionId,
	UseViewerBuildingBlocksActionsOptions,
	ViewerBuildingBlockActions,
	ViewerMode,
} from 'ooxml-ui/pptx';

import type {
	SlidesPaneLayoutGalleryRequest,
	SlidesPaneProps,
} from '../viewer/components/slides-pane-props';
import type { ViewerStateBag } from '../viewer/state/create-viewer-state-types';
import type { PowerPointViewerProps, ViewerLoadDetail } from '../viewer/types';

/** The viewer options the building blocks take (they mount no chrome of their own). */
type ForwardedProps =
	| 'fileName'
	| 'filePath'
	| 'initialSlide'
	| 'locale'
	| 'googleFonts'
	| 'listConversionScope'
	| 'allowExternalContent'
	| 'limits'
	| 'authorName'
	| 'collaboration'
	| 'fitPadding'
	| 'maxFitScale';

export interface ViewerBuildingBlocksOptions
	extends Pick<PowerPointViewerProps, ForwardedProps>, UseViewerBuildingBlocksActionsOptions {
	/** The presentation bytes, or null while no file is open. */
	content?: Uint8Array | ArrayBuffer | null | undefined;
	/** Whether editing is enabled. Defaults to false (view-only). */
	canEdit?: boolean | undefined;
	/** Whether the built-in AutoRecover snapshots run. Off unless the host opts in. */
	autosave?: boolean | undefined;
	/** As React's `userName`: the author of comments (`authorName` wins when both are set). */
	userName?: string | undefined;
	/** Toolbar buttons to hide from the compact toolbar (`undo`, `redo`, `insert`, `zoom`, `navigation`). */
	hiddenActions?: readonly ToolbarActionId[] | undefined;
	onDirtyChange?: ((dirty: boolean) => void) | undefined;
	onSlideChange?: ((index: number) => void) | undefined;
	onLoad?: ((detail: ViewerLoadDetail) => void) | undefined;
	onError?: ((message: string) => void) | undefined;
	onContentChange?: ((content: Uint8Array) => void) | undefined;
	onSelectionChange?: ((elementIds: string[]) => void) | undefined;
	/**
	 * Fired by the "Layout" entry of a slide's context menu in the slides pane, after that slide is
	 * made active; the host owns rendering the Layout gallery.
	 */
	onOpenLayoutGallery?: ((request: SlidesPaneLayoutGalleryRequest) => void) | undefined;
}

/** The state a custom shell renders around the blocks (a status line, Save enablement). */
export interface ViewerBuildingBlocksState {
	/** The effective mode: `preview` while the blocks may not edit. */
	readonly mode: ViewerMode;
	readonly loading: boolean;
	readonly error: string | null;
	/** Zero-based active slide, -1 before a deck loads. */
	readonly slideIndex: number;
	readonly slideCount: number;
	readonly canUndo: boolean;
	readonly canRedo: boolean;
	readonly dirty: boolean;
}

/** The DOM handles the canvas block reports back (the viewer state measures and listens through them). */
export interface ViewerBuildingBlocksCanvasHandles {
	setRoot(element: HTMLDivElement | undefined): void;
	setStageHolder(element: HTMLDivElement | undefined): void;
	setViewportSize(size: { width: number; height: number }): void;
}

/**
 * The Svelte building blocks: create it once in the component that renders the blocks (it builds
 * the viewer state, which registers effects and context) and hand it to `BuildingBlocksCanvas`,
 * `BuildingBlocksToolbar` and `BuildingBlocksSlidesPane` as `blocks`. Every member is a live
 * runes-backed getter, so reading it in a template or `$derived` tracks the editor.
 */
export interface ViewerBuildingBlocks {
	/** Host-owned workflows (pickers, dialogs, Find / Replace, layouts), shared by every binding. */
	readonly actions: ViewerBuildingBlockActions;
	readonly state: ViewerBuildingBlocksState;
	/** The compact toolbar's groups and controls (existing dictionary keys, `describeBuildingBlockToolbar`). */
	readonly toolbar: BuildingBlockToolbarGroup[];
	/** Run a compact toolbar control (what the toolbar block's buttons call). */
	runToolbarControl(id: BuildingBlockToolbarControlId): void;
	/** Props for the slides pane block (`createSlidesPaneProps` over the viewer state). */
	readonly slidesPaneProps: SlidesPaneProps;
	/** The find matches of the active slide the canvas marks in the text, by element. */
	readonly findHighlights: SlideFindHighlights;
	/** Translate a dictionary key with the blocks' locale. */
	t(key: string): string;
	/** The viewer state behind the blocks, for the imperative API (`deck`, `editingApi`, ...). */
	readonly viewerState: ViewerStateBag;
	/** The shared session, for hosts that register further action groups. */
	readonly session: PptxBuildingBlockSession;
	/** What the canvas block reports back; not for hosts. */
	readonly canvas: ViewerBuildingBlocksCanvasHandles;
	/** Tear down the viewer state; also runs when the creating component is destroyed. */
	destroy(): void;
}
