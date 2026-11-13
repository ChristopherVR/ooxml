/**
 * The public shape of `actions` on every binding's building blocks (React
 * `useViewerBuildingBlocks().actions`, vanilla `createViewerBuildingBlocks().actions`, and the Vue,
 * Svelte and Angular ports). One type, so a host's pickers and dialogs work against any binding.
 *
 * @module render/building-blocks/actions-types
 */
import type { PptxLayoutOption, SmartArtLayout } from 'ooxml-core/pptx';

import type {
	BuildingBlockActionId,
	BuildingBlockActionState,
	BuildingBlockResult,
	HostFileLimits,
} from '../building-block-actions';
import type { FindResult } from '../find-replace';
import type { FindReplaceOptions } from '../find-replace-session';
import type { InsertChartKind, InsertChartOptions } from '../insert-chart';
import type { BuildingBlockActionGroupMap } from './action-groups';
import type { HyperlinkEditData } from './hyperlink-edit-data';
import type { BuildingBlockRequest } from './requests';

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

export interface BuildingBlockHyperlinkActions {
	/** The link on the selected element, as editor values; null when there is none. */
	current: HyperlinkEditData | null;
	/** For a slide-target picker: how many slides can be linked to. */
	slideCount: number;
	apply: (data: HyperlinkEditData) => BuildingBlockResult;
	remove: () => BuildingBlockResult;
}

export interface BuildingBlockEquationActions {
	/** The selected element's equation (OMML object), null when it has none. */
	current: Record<string, unknown> | null;
	insert: (omml: Record<string, unknown>) => BuildingBlockResult;
	/** Replace the selected element's equation. */
	update: (omml: Record<string, unknown>) => BuildingBlockResult;
}

/** The workflows every binding exposes; registered groups add their own keys beside these. */
export interface ViewerBuildingBlockCoreActions {
	/** What each workflow can do right now and whether its toolbar control needs host UI. */
	availability: Record<BuildingBlockActionId, BuildingBlockActionState>;
	/**
	 * Toolbar handlers that open a dialog or pane the building blocks do not mount (set internal
	 * state, show nothing). Hide them with `hiddenActions` or render the dialog yourself.
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
	hyperlink: BuildingBlockHyperlinkActions;
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
	equation: BuildingBlockEquationActions;
	findReplace: BuildingBlockFindReplace;
	layout: BuildingBlockLayoutActions;
}

/** `actions` as every binding returns it: the core workflows plus every registered group. */
export type ViewerBuildingBlockActions = ViewerBuildingBlockCoreActions &
	Partial<BuildingBlockActionGroupMap>;
