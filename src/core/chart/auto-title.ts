/**
 * The text a chart's title shows, including Office's AUTO title (COM-verified in PowerPoint).
 *
 * A `c:title` with no `c:tx` is not an empty title: Office shows the only series' name for a
 * single-series chart (a pie titled "Sales") and the literal "Chart Title" otherwise. The
 * Office 2016+ extended (`cx:`) chart types have no "lone series name" convention: an empty
 * `cx:title` always renders "Chart Title", even for a single-series waterfall or funnel.
 */

/** Office's placeholder for an auto title on a multi-series chart. */
export const CHART_AUTO_TITLE = 'Chart Title';

const CHARTEX_ONLY_TYPES: ReadonlySet<string> = new Set([
	'waterfall',
	'funnel',
	'treemap',
	'sunburst',
	'boxWhisker',
	'histogram',
	'regionMap',
]);

/** Whether `chartType` is one of the `cx:` (chartex) families. */
export function isChartexOnlyType(chartType: string): boolean {
	return CHARTEX_ONLY_TYPES.has(chartType);
}

export interface AutoTitleInput {
	chartType: string;
	/** Series names, in order (an empty string for an unnamed series). */
	seriesNames: readonly string[];
	/** The explicit title, when the chart authors one. */
	title?: string | undefined;
	/** `style.hasTitle` as authored; `undefined` when the chart does not say. */
	hasTitle?: boolean | undefined;
	/** `c:autoTitleDeleted`. */
	autoTitleDeleted?: boolean | undefined;
}

/**
 * Whether a classic chart with NO `c:title` element still gets an auto title: exactly one named
 * series, no authored toggle, `c:autoTitleDeleted` unset and not a chartex family.
 */
export function impliesAutoTitle(input: AutoTitleInput): boolean {
	return !(
		input.hasTitle !== undefined ||
		input.autoTitleDeleted === true ||
		isChartexOnlyType(input.chartType) ||
		input.seriesNames.length !== 1 ||
		!input.seriesNames[0]
	);
}

/** The title text to draw, or `undefined` when the chart shows no title. */
export function resolveAutoTitleText(input: AutoTitleInput): string | undefined {
	if (!input.hasTitle) return undefined;
	if (input.title) return input.title;
	if (!isChartexOnlyType(input.chartType) && input.seriesNames.length === 1)
		return input.seriesNames[0] || undefined;
	return input.seriesNames.length > 0 ? CHART_AUTO_TITLE : undefined;
}
