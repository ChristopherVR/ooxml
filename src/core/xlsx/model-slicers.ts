import type { DrawingAnchor } from './model';

/** Default slicer button height in EMU (`CT_Slicer/@rowHeight`), 26.4 px at 96 dpi. */
export const SLICER_DEFAULT_ROW_HEIGHT = 251883;

/** A slicer's data source: a table column or one PivotTable field. */
export type SlicerSource =
	| {
			kind: 'table';
			/** `table/@id` of the filtered table. */
			tableId: number;
			/** `tableColumn/@id` (1-based in Excel's files) of the filtered column. */
			column: number;
	  }
	| {
			kind: 'pivot';
			/** `x14:pivotCacheDefinition/@pivotCacheId` linking the slicer to its pivot cache. */
			pivotCacheId: number;
			/** The connected PivotTables: sheet id (`tabId`) and name. */
			pivots: { tabId: number; name: string }[];
	  };

/**
 * A slicer cache (`slicerCacheDefinition`): what the slicer filters. The selection itself lives
 * with the filtered object (`TableColumn.filter`, `PivotTable.itemFilters`), exactly as in Excel.
 */
export interface SlicerCache {
	/** `Slicer_Category`: the cache name and the matching defined name. */
	name: string;
	uid?: string;
	/** The source column or field name. */
	sourceName: string;
	source: SlicerSource;
	sortOrder?: 'ascending' | 'descending';
	/** `showMissing="0"` / the x15 hide flag: leave out items without data. */
	hideNoData?: boolean;
	/** `crossFilter` of the tabular data: how items without data in other slicers show. */
	crossFilter?: 'showItemsWithDataAtTop' | 'showItemsWithNoData' | 'none';
	/** The cache part as read; a save patches it so unmodelled markup survives. */
	sourceXml?: string;
	/** The source package part, reused on save. */
	partName?: string;
}

/** A slicer (`slicer` element of a `slicers` part), drawn from its drawing anchor. */
export interface Slicer {
	name: string;
	uid?: string;
	/** Name of the slicer cache. */
	cache: string;
	caption?: string;
	/** Built-in (`SlicerStyleLight1`) or custom slicer style name. */
	style?: string;
	/** Item columns (1 or more). */
	columnCount?: number;
	/** Button height in EMU. */
	rowHeight?: number;
	/** Show the caption header (default true). */
	showCaption?: boolean;
	startItem?: number;
	lockedPosition?: boolean;
	/** `level` for hierarchical sources; always 0 for the supported tabular sources. */
	level?: number;
	/** Extension list as read, written back verbatim. */
	extLstXml?: string;
	/** The `slicers` part this slicer was read from. */
	partName?: string;
	/** Sheet that draws it (index), set by the reader and the insert command. */
	sheet: number;
}

export type TimelineLevel = 'years' | 'quarters' | 'months' | 'days';

/** A timeline cache (`timelineCacheDefinition`): the date field it filters and its range. */
export interface TimelineCache {
	name: string;
	uid?: string;
	sourceName: string;
	pivotCacheId: number;
	pivots: { tabId: number; name: string }[];
	/** The selected range as ISO dates; absent when everything is selected. */
	selection?: { start: string; end: string };
	/** The data bounds as ISO dates. */
	bounds: { start: string; end: string };
	filterType?: string;
	/** `state/@minimalRefreshVersion` and `lastRefreshVersion`, kept for writing. */
	minimalRefreshVersion?: number;
	lastRefreshVersion?: number;
	/** The cache part as read; a save patches it so unmodelled markup survives. */
	sourceXml?: string;
	partName?: string;
}

/** A timeline (`timeline` element of a `timelines` part). */
export interface Timeline {
	name: string;
	uid?: string;
	cache: string;
	caption?: string;
	showHeader?: boolean;
	showSelectionLabel?: boolean;
	showTimeLevel?: boolean;
	showHorizontalScrollbar?: boolean;
	/** Visible period level (`level`: 0 years, 1 quarters, 2 months, 3 days). */
	level: number;
	selectionLevel: number;
	/** First visible date, ISO. */
	scrollPosition?: string;
	style?: string;
	extLstXml?: string;
	partName?: string;
	sheet: number;
}

/** A slicer or timeline frame on a sheet; the anchor XML is kept for lossless saves. */
interface SlicerFrame<K extends 'slicer' | 'timeline'> {
	kind: K;
	anchor: DrawingAnchor;
	/** Matches `Slicer.name` / `Timeline.name`. */
	name: string;
	/** The `xdr:*Anchor` element as read (with its `mc:AlternateContent` fallback). */
	sourceXml: string;
}

export type SlicerFrameObject = SlicerFrame<'slicer'> | SlicerFrame<'timeline'>;

/** Everything the workbook's slicers and timelines need besides the drawn frames. */
export interface SlicerBook {
	/** Set by edits: the slicer and timeline parts are then rewritten from the model. */
	dirty?: boolean;
	/** Source parts whose slicers or caches were all deleted; a save leaves them out. */
	removedParts?: string[];
	caches: SlicerCache[];
	slicers: Slicer[];
	timelineCaches: TimelineCache[];
	timelines: Timeline[];
}

/** Value-list filter of a table column, as a slicer writes it. */
export interface TableColumnFilter {
	/** Display texts that stay visible (`<filter val>`). */
	values: string[];
	/** Blank cells stay visible (`<filter blank="1">`). */
	blank?: boolean;
}
