// Custom Views (View > Custom Views): `workbook.customWorkbookViews` names a view, and every
// sheet holds a `customSheetView` with the same guid that records how that sheet looked. Hidden
// rows and columns, the filter range and the print area/titles live in hidden defined names
// `Z_<GUID>_.wvu.Rows`, `.Cols`, `.FilterData`, `.PrintArea` and `.PrintTitles` (kept in
// `Workbook.definedNames`, one set per view and sheet).
import type { AutoFilter, PageSetup, PrintOptions, SheetView } from './model';

/** One named view of the workbook (`customWorkbookView`). */
export interface CustomWorkbookView {
	/** `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`. */
	guid: string;
	name: string;
	/** Printer settings are part of the view (default true). */
	includePrintSettings: boolean;
	/** Hidden rows, columns and filter settings are part of the view (default true). */
	includeHiddenRowCol: boolean;
	/** Window position, size and the like, and `activeSheetId`, kept as read. */
	attrs: Record<string, string>;
}

/** How one sheet looked in a custom view (`customSheetView`). */
export interface CustomSheetView {
	guid: string;
	/** Zoom, grid lines, headings, panes and selection. */
	view: SheetView;
	/** The names store the hidden rows and columns; these flags say they exist. */
	hiddenRows?: true;
	hiddenColumns?: true;
	/** The auto-filter and its criteria, when the sheet was filtered. */
	filter?: AutoFilter;
	/** The auto-filter arrows were showing (`showAutoFilter`). */
	showAutoFilter?: true;
	pageSetup?: PageSetup;
	printOptions?: PrintOptions;
	/** Other attributes of the element (`showPageBreaks`, `state`, `colorId`...), kept as read. */
	attrs: Record<string, string>;
	/** Children the model does not interpret (`rowBreaks`, `colBreaks`, `extLst`), kept as read. */
	extra: string[];
}
