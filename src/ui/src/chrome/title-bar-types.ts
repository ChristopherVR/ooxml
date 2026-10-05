export type OfficeTitleBarPlacement = 'titleBar' | 'belowRibbon';
export type OfficeTitleBarTone = 'idle' | 'saving' | 'error';

/** The AutoSave switch. The host owns the state; a change is only a request. */
export interface OfficeTitleBarAutosave {
	enabled: boolean;
	/** Default true. False renders the switch inert. */
	available?: boolean | undefined;
	/** "AutoSave". */
	label: string;
	/** "On" / "Off", shown after the switch. */
	stateLabel: string;
	/** Accessible name of the switch. */
	toggleLabel: string;
	/** ScreenTip of the switch; defaults to `toggleLabel`. */
	title?: string | undefined;
}

/** One Quick Access Toolbar button; every string arrives translated from the host. */
export interface OfficeQuickAccessItem {
	id: string;
	/** Registered glyph name (see `registerIcon`). */
	icon: string;
	/** Accessible name, and the visible text when labels show. */
	label: string;
	/** ScreenTip; absent leaves the button without a `title`. */
	title?: string | undefined;
	disabled?: boolean | undefined;
}

/** One command the title-bar search can run. `category` shows on the right of its row. */
export interface OfficeTitleBarCommand {
	id: string;
	label: string;
	category?: string | undefined;
}

/** The centred command search. Every string arrives translated from the host. */
export interface OfficeTitleBarSearch {
	placeholder: string;
	label: string;
	/** Heading over the matches and the live-region text when there are some. */
	heading: string;
	/** Row and live-region text when nothing matches. */
	empty: string;
	/** Searched by label when `match` is absent. */
	commands?: readonly OfficeTitleBarCommand[] | undefined;
	/** Host ranking (synonyms, keywords); replaces the default label search. */
	match?: ((query: string) => readonly OfficeTitleBarCommand[]) | undefined;
	/** Text of the trailing "search the document" row; absent drops that row. */
	content?: ((query: string) => string) | undefined;
	/** Glyph of the content row; defaults to `search`. */
	contentIcon?: string | undefined;
}

/** Detail of the search request: a chosen command, or a document search when `command` is absent. */
export interface OfficeTitleBarSearchDetail {
	query: string;
	command?: string;
}

/** Everything the bar shows, already translated. Absent parts are not rendered. */
export interface OfficeTitleBarState {
	/** One or two letters in the product-coloured app mark; absent hides the mark. */
	appMark?: string | undefined;
	fileName: string;
	/** Save-location or save-progress text after the name. */
	status?: string | undefined;
	tone?: OfficeTitleBarTone | undefined;
	autosave?: OfficeTitleBarAutosave | undefined;
	quickAccess?:
		| { label: string; items: readonly OfficeQuickAccessItem[]; showLabels?: boolean | undefined }
		| undefined;
	search?: OfficeTitleBarSearch | undefined;
}

/** The dropdown shows at most this many commands so it never outgrows its box. */
export const OFFICE_TITLE_BAR_SEARCH_LIMIT = 8;

/** The default search: commands whose label contains the query. */
export function matchTitleBarCommands(
	search: OfficeTitleBarSearch,
	query: string,
): readonly OfficeTitleBarCommand[] {
	if (search.match) return search.match(query);
	const needle = query.trim().toLowerCase();
	return (search.commands ?? []).filter((entry) => entry.label.toLowerCase().includes(needle));
}
