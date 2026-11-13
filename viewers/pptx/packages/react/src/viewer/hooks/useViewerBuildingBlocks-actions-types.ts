import type {
	BuildingBlockActionId,
	BuildingBlockActionState,
	BuildingBlockResult,
	FindReplaceOptions,
	FindResult,
	HostFileLimits,
	ShowDialogActionId,
	ShowDialogRequest,
	InsertChartKind,
	InsertChartOptions,
} from 'ooxml-ui/pptx';
import type { PptxLayoutOption, SmartArtLayout } from 'pptx-viewer-core';

import type { HyperlinkEditData } from '../components/hyperlink-edit-types';
import type { FileDialogActions } from './useViewerBuildingBlocks-file-dialog-actions';
import type { ShowDialogActions } from './useViewerBuildingBlocks-show-dialog-actions';

/**
 * A toolbar control or canvas gesture asked for UI the building blocks never mount (a file picker,
 * the hyperlink, SmartArt or equation dialog). The host renders it and finishes the workflow with
 * the matching operation on {@link ViewerBuildingBlockActions}, or calls `cancel()`.
 */
export type BuildingBlockRequest =
	| { kind: 'image' }
	| { kind: 'media' }
	| { kind: 'hyperlink' }
	| { kind: 'smartArt' }
	| { kind: 'chart' }
	/** Open password, Mark as Final: render your own UI from `actions.passwordProtection`. */
	| { kind: 'passwordProtection' }
	/** Render your own editor from `actions.documentProperties`. */
	| { kind: 'documentProperties' }
	/** Render your own options from `actions.fontEmbedding`. */
	| { kind: 'fontEmbedding' }
	/** Render your own list from `actions.digitalSignatures`. */
	| { kind: 'digitalSignatures' }
	/** `existingOmml` is the equation being edited (a double-click), null for a new one. */
	| { kind: 'equation'; existingOmml: Record<string, unknown> | null }
	/** Set Up Slide Show, Custom Shows, or the shortcut help opening or closing. */
	| ShowDialogRequest;

export interface UseViewerBuildingBlocksActionsOptions {
	/**
	 * Called when a toolbar control or the canvas raises a {@link BuildingBlockRequest}. The same
	 * request is also readable as `actions.pending` until it is completed or cancelled.
	 */
	onRequest?: (request: BuildingBlockRequest) => void;
	/** Ceilings for host-picked files; the defaults are 25 MiB (image) and 100 MiB (audio, video). */
	fileLimits?: HostFileLimits;
}

export interface BuildingBlockFindReplace {
	/** Whether the find bar is open (also toggled by Ctrl+F / Ctrl+H). The host renders the bar. */
	open: boolean;
	setOpen: (open: boolean) => void;
	query: string;
	replacement: string;
	options: FindReplaceOptions;
	matches: readonly FindResult[];
	/** Zero-based index of the current match, -1 when there is none. */
	currentIndex: number;
	currentMatch: FindResult | null;
	setQuery: (query: string) => void;
	setReplacement: (replacement: string) => void;
	setOptions: (options: Partial<FindReplaceOptions>) => void;
	/** Search again with the current query and options. */
	find: () => void;
	/** Go to the next / previous match; the viewer shows its slide and selects the element. */
	next: () => void;
	previous: () => void;
	/** Replace the current match / every match. Edit permission, validation and history apply. */
	replaceCurrent: () => BuildingBlockResult;
	replaceAll: () => BuildingBlockResult;
}

export interface BuildingBlockLayoutActions {
	/** The active slide's layout, when it records one the deck lists. */
	current: PptxLayoutOption | undefined;
	currentPath: string | undefined;
	/** The layouts the active slide may switch to (its master's). */
	available: PptxLayoutOption[];
	apply: (path: string) => Promise<BuildingBlockResult>;
	/** Re-apply the slide's own layout (PowerPoint's Reset). */
	reset: () => Promise<BuildingBlockResult>;
}

export interface ViewerBuildingBlockActions extends FileDialogActions, ShowDialogActions {
	/** What each workflow can do right now and whether its toolbar control needs host UI. */
	availability: Record<BuildingBlockActionId | ShowDialogActionId, BuildingBlockActionState>;
	/**
	 * `ToolbarProps` handlers that open a dialog or pane the building blocks do not mount (set
	 * internal state, show nothing). Hide them with `hiddenActions` or render the dialog yourself.
	 */
	unmountedDialogActions: readonly string[];
	/** The request raised by the toolbar or canvas that the host has not completed yet. */
	pending: BuildingBlockRequest | null;
	/** Abandon the pending request without changing the deck. */
	cancel: () => void;
	/** Insert a picture chosen by the host's own picker, centred and fitted to the slide. */
	insertImage: (file: File) => Promise<BuildingBlockResult>;
	/** Insert an audio or video file chosen by the host's own picker. */
	insertMedia: (file: File) => Promise<BuildingBlockResult>;
	hyperlink: {
		/** The link on the selected element, as editor values; null when there is none. */
		current: HyperlinkEditData | null;
		/** For a slide-target picker: how many slides can be linked to. */
		slideCount: number;
		apply: (data: HyperlinkEditData) => BuildingBlockResult;
		remove: () => BuildingBlockResult;
	};
	smartArt: {
		insert: (layout: SmartArtLayout, items: string[]) => BuildingBlockResult;
	};
	chart: {
		/**
		 * Insert a chart of `kind` (the Insert Chart dialog's entries, for example 'column' or
		 * 'pie') with optional title, grouping, legend and data (otherwise sample data).
		 */
		insert: (kind: InsertChartKind, options?: InsertChartOptions) => BuildingBlockResult;
	};
	equation: {
		/** The selected element's equation (OMML object), null when it has none. */
		current: Record<string, unknown> | null;
		insert: (omml: Record<string, unknown>) => BuildingBlockResult;
		/** Replace the selected element's equation. */
		update: (omml: Record<string, unknown>) => BuildingBlockResult;
	};
	findReplace: BuildingBlockFindReplace;
	layout: BuildingBlockLayoutActions;
}
