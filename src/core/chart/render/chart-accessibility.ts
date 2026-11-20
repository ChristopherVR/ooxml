// What a chart offers assistive technology: its accessible name and a readout of its data.
//
// A chart is a picture, so a screen reader gets only the name and description its host provides.
// The name is the author's alt text (the graphic frame's `cNvPr/@descr`, Word's `wp:docPr/@descr`),
// which is the text the author wrote for exactly this purpose; the chart's visible title is the
// fallback when there is none. The readout lists each series with its categories and values, so
// the numbers behind the picture are reachable without a sighted reading of the plot.
//
// The readout carries data and punctuation only, never a word of prose, so it needs no translation
// and no per-product dictionary. Hosts put it in the SVG `<desc>` (`renderChartSvg`) or, where the
// frame around the SVG is itself the image, in an `aria-describedby` target.

/** Series and categories a readout needs; both `ChartViewModel` and `PptxChartData` satisfy it. */
export interface ChartDataReadout {
	categories?: readonly (string | number | null | undefined)[];
	/** Absent on a chart whose part could not be read; there is then nothing to read out. */
	series?: readonly {
		name?: string;
		/** A `ChartSummary` keeps the categories per series; a view model keeps one list for all. */
		categories?: readonly (string | number | null | undefined)[];
		values: readonly (number | null | undefined)[];
		hidden?: boolean;
	}[];
}

/** Series a readout lists, and points per series, before it ends in an ellipsis. */
const MAX_SERIES = 12;
const MAX_POINTS = 24;

/**
 * The chart's accessible name: the author's alt text, else the chart's own title. Empty or
 * whitespace-only text counts as absent, as it does in Office's Alt Text pane.
 */
export function chartAccessibleName(
	chart: { title?: string | undefined },
	altText?: string | undefined,
): string | undefined {
	return altText?.trim() || chart.title?.trim() || undefined;
}

/** A value as a reader should hear it: no exponent noise, no trailing float dust. */
function value(input: number): string {
	return String(Number(input.toPrecision(12)));
}

/**
 * The chart's data as one clause per series: `Revenue: Q1 12; Q2 18. Cost: Q1 5; Q2 6.` Series with
 * no plotted value, and series the Selection pane hid, are left out. Empty when there is no data.
 */
export function chartDataDescription(chart: ChartDataReadout): string {
	const clauses: string[] = [];
	let more = false;
	for (const series of chart.series ?? []) {
		if (series.hidden) continue;
		if (clauses.length === MAX_SERIES) {
			more = true;
			break;
		}
		const points: string[] = [];
		const labels = series.categories ?? chart.categories;
		for (const [at, raw] of series.values.entries()) {
			if (raw === null || raw === undefined || !Number.isFinite(raw)) continue;
			if (points.length === MAX_POINTS) {
				points.push('…');
				break;
			}
			const label = labels?.[at];
			const category = label === null || label === undefined ? '' : String(label).trim();
			points.push(category ? `${category} ${value(raw)}` : value(raw));
		}
		if (!points.length) continue;
		const name = series.name?.trim();
		clauses.push(`${name ? `${name}: ` : ''}${points.join('; ')}`);
	}
	if (!clauses.length) return '';
	return `${clauses.join('. ')}.${more ? ' …' : ''}`;
}
