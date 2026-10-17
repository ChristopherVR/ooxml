// What a chart data label says, independent of the chart model it is read from.
//
// `c:dLbls` carries `showVal`, `showCatName`, `showSerName`, `showPercent`, `showBubbleSize` and
// `c:separator`; Office combines whichever are set. The flags live in three places and override
// each other point first, then series, then chart group. Moved from the PowerPoint renderer
// (`ooxml-ui/pptx` `chart-data-label-text.ts`), which now calls these two functions with its own
// level sources; the classic and chart-extension models both feed them.
import { formatChartNumberWithColor } from './number-format';

/** The content switches one cascade level (a `c:dLbl`, a series' or a group's `c:dLbls`) writes. */
export interface DataLabelFlagLevel {
	showValue?: boolean | undefined;
	showCategory?: boolean | undefined;
	showSeriesName?: boolean | undefined;
	showPercent?: boolean | undefined;
	showBubbleSize?: boolean | undefined;
	separator?: string | undefined;
}

/** The levels a point's label reads, highest priority first. */
export interface DataLabelFlagLevels {
	point?: DataLabelFlagLevel | undefined;
	series?: DataLabelFlagLevel | undefined;
	chart?: DataLabelFlagLevel | undefined;
	/**
	 * No level here declares anything, but another series writes its own labels: this series then
	 * shows none, rather than falling back to the value.
	 */
	siblingOwnsLabels?: boolean | undefined;
}

/** The label content flags once the cascade has been applied. */
export interface DataLabelFlags {
	showValue: boolean;
	showCategory: boolean;
	showSeriesName: boolean;
	showPercent: boolean;
	showBubbleSize: boolean;
	/** Text joining the enabled components. */
	separator: string;
}

/** Office's default separator between combined label components. */
export const DATA_LABEL_SEPARATOR = ', ';

/** Office's default when a label shows the category name and the percentage: a line break. */
export const DATA_LABEL_CATEGORY_PERCENT_SEPARATOR = '\n';

const FLAG_KEYS = [
	'showValue',
	'showCategory',
	'showSeriesName',
	'showPercent',
	'showBubbleSize',
] as const;

/**
 * Applies the point > series > group cascade to the content flags.
 *
 * A flag counts as declared when any level writes it, even as `0`: Office writes an all-zero
 * group level on every chart it authors, and that means "no labels", never "fall back to the
 * value". When nothing is declared anywhere the label shows the value only.
 */
export function resolveDataLabelFlags(levels: DataLabelFlagLevels): DataLabelFlags {
	const { point, series, chart } = levels;
	const picked = {} as Record<(typeof FLAG_KEYS)[number], boolean | undefined>;
	for (const key of FLAG_KEYS) picked[key] = point?.[key] ?? series?.[key] ?? chart?.[key];
	const anyDeclared = FLAG_KEYS.some((key) => picked[key] !== undefined);
	return {
		showValue: anyDeclared ? picked.showValue === true : levels.siblingOwnsLabels !== true,
		showCategory: picked.showCategory === true,
		showSeriesName: picked.showSeriesName === true,
		showPercent: picked.showPercent === true,
		showBubbleSize: picked.showBubbleSize === true,
		separator:
			point?.separator ??
			series?.separator ??
			chart?.separator ??
			(picked.showCategory === true && picked.showPercent === true
				? DATA_LABEL_CATEGORY_PERCENT_SEPARATOR
				: DATA_LABEL_SEPARATOR),
	};
}

/** What one label is built from. */
export interface DataLabelTextInput {
	flags: DataLabelFlags;
	seriesName?: string | undefined;
	category?: string | undefined;
	value: number;
	/** Number format code of the label (point, series, group or source cells). */
	numberFormat?: string | undefined;
	/** Denominator of `showPercent`: the sum of the series' absolute values on a pie. */
	percentBase: number;
	bubbleSize?: number | undefined;
	/** Value text when the format code is absent, `General` or unsupported. Default `String`. */
	formatValue?: ((value: number, numberFormat: string | undefined) => string) | undefined;
	/** Bubble size text. Default `String`. */
	formatBubbleSize?: ((size: number) => string) | undefined;
}

/** A built data label: its text and the `[Red]`/`[Blue]` colour of its number format, if any. */
export interface DataLabelText {
	text: string;
	/** Only set when the label is the formatted value alone (a joined label cannot tint one part). */
	color?: string | undefined;
}

/**
 * The text of one label in Office's component order (series name, category, value, percentage,
 * bubble size), or `undefined` when it resolves to nothing.
 */
export function composeDataLabelText(input: DataLabelTextInput): DataLabelText | undefined {
	const { flags, value, numberFormat } = input;
	const parts: string[] = [];
	let valueColor: string | undefined;
	if (flags.showSeriesName && input.seriesName) parts.push(input.seriesName);
	if (flags.showCategory && input.category) parts.push(input.category);
	if (flags.showValue) {
		const formatted = formatChartNumberWithColor(value, numberFormat);
		parts.push(formatted?.text ?? (input.formatValue ?? String)(value, numberFormat));
		valueColor = formatted?.color;
	}
	if (flags.showPercent) {
		const base = input.percentBase;
		parts.push(base > 0 ? `${Math.round((Math.abs(value) / base) * 100)}%` : '0%');
	}
	if (flags.showBubbleSize && input.bubbleSize !== undefined)
		parts.push((input.formatBubbleSize ?? String)(input.bubbleSize));
	const text = parts.join(flags.separator);
	if (text.length === 0) return undefined;
	return { text, color: parts.length === 1 && flags.showValue ? valueColor : undefined };
}
