// Edit operations over the neutral extended chart model (`ChartExSpace`): title, legend, data
// labels, series names and fills, and the family options (waterfall totals, histogram binning,
// treemap labels, box and whisker statistics). Pure functions returning a new model. They follow
// PowerPoint's ChartEx save (pptx `runtime/chart-cx-update*.ts`), the checklist of what PowerPoint
// writes; where this module differs it says so.
import type { DrawingFill } from '../../drawingml/types';
import type {
	ChartExBinning,
	ChartExLayoutProperties,
	ChartExSeries,
	ChartExSeriesVisibility,
	ChartExSpace,
} from '../ex/model';
import { chartTextBodyFromXml } from '../write-shape';
import { escapeText } from '../write-util';
import { retextChartTextBody } from './title';

/** A copy of `space` changed by `recipe`; `space` itself is never mutated. */
export function editChartExSpace(
	space: ChartExSpace,
	recipe: (draft: ChartExSpace) => void,
): ChartExSpace {
	const draft = structuredClone(space);
	recipe(draft);
	return draft;
}

/** The series at `index` (compared as a position, never used as a property key). */
export function chartExSeriesAt(space: ChartExSpace, index: number): ChartExSeries {
	const series =
		typeof index === 'number' && Number.isSafeInteger(index)
			? space.plotArea.series.find((_, position) => position === index)
			: undefined;
	if (!series) throw new RangeError(`No chart series at index ${String(index)}`);
	return series;
}

/**
 * Sets the chart title text, or removes the title (`undefined`). An existing rich title keeps its
 * formatting and its first run takes the text; a title linked to a cell (`cx:txData`) or a new one
 * becomes one rich run, as PowerPoint writes it (with the `a:bodyPr` the schema requires, which
 * PowerPoint's builder leaves out). Semantics: PowerPoint.
 */
export function setChartExTitle(space: ChartExSpace, text: string | undefined): ChartExSpace {
	return editChartExSpace(space, (draft) => {
		if (text === undefined) {
			delete draft.title;
			return;
		}
		const rich = draft.title?.tx?.rich;
		const body =
			(rich && retextChartTextBody(rich, text)) ??
			chartTextBodyFromXml(`<a:bodyPr/><a:p><a:r><a:t>${escapeText(text)}</a:t></a:r></a:p>`);
		draft.title = { ...draft.title, tx: { rich: body } };
	});
}

/** ChartEx legend positions are `ST_SidePos`; the classic corner `tr` folds onto the right. */
const sidePosition = (position: string | undefined) =>
	position === 't' || position === 'b' || position === 'l' ? position : 'r';

/**
 * Shows, moves or hides the legend. A shown legend keeps its alignment, overlay and formatting; a
 * new one is centred and not overlaid. Semantics: PowerPoint.
 */
export function setChartExLegend(
	space: ChartExSpace,
	options: { visible: boolean; position?: string },
): ChartExSpace {
	return editChartExSpace(space, (draft) => {
		if (!options.visible) {
			delete draft.legend;
			return;
		}
		if (draft.legend) {
			if (options.position !== undefined) draft.legend.position = sidePosition(options.position);
			return;
		}
		draft.legend = { position: sidePosition(options.position), align: 'ctr', overlay: false };
	});
}

/**
 * Shows or hides the data labels of every series. Hiding removes `cx:dataLabels`; showing keeps
 * existing labels and gives the others category name and value. Semantics: PowerPoint.
 */
export function setChartExDataLabelsVisible(space: ChartExSpace, visible: boolean): ChartExSpace {
	return editChartExSpace(space, (draft) => {
		for (const series of draft.plotArea.series) {
			if (!visible) delete series.dataLabels;
			else
				series.dataLabels ??= {
					visibility: { seriesName: false, categoryName: true, value: true },
					labels: [],
					hidden: [],
				};
		}
	});
}

/**
 * Renames the series at `index`. A name linked to a cell keeps its formula with the new cached
 * text (pass `keepReference: false` to drop the link); a rich or missing name becomes cached text.
 * Semantics: PowerPoint.
 */
export function renameChartExSeries(
	space: ChartExSpace,
	index: number,
	name: string,
	options: { keepReference?: boolean } = {},
): ChartExSpace {
	return editChartExSpace(space, (draft) => {
		const series = chartExSeriesAt(draft, index);
		const formula = options.keepReference === false ? undefined : series.tx?.formula;
		series.tx = formula ? { formula, value: name } : { value: name };
	});
}

/** Sets (or, with `undefined`, removes) the fill of the series at `index`; the outline stays. */
export function setChartExSeriesFill(
	space: ChartExSpace,
	index: number,
	fill: DrawingFill | undefined,
): ChartExSpace {
	return editChartExSpace(space, (draft) => {
		const series = chartExSeriesAt(draft, index);
		if (fill) series.spPr = { ...series.spPr, fill };
		else if (series.spPr) delete series.spPr.fill;
	});
}

function editLayout(
	space: ChartExSpace,
	index: number,
	recipe: (layout: ChartExLayoutProperties, series: ChartExSeries) => void,
): ChartExSpace {
	return editChartExSpace(space, (draft) => {
		const series = chartExSeriesAt(draft, index);
		series.layoutProperties ??= {};
		recipe(series.layoutProperties, series);
	});
}

function setVisibility(layout: ChartExLayoutProperties, flags: ChartExSeriesVisibility): void {
	const merged: ChartExSeriesVisibility = { ...layout.visibility };
	for (const [key, value] of Object.entries(flags) as [keyof ChartExSeriesVisibility, unknown][])
		if (typeof value === 'boolean') merged[key] = value;
	layout.visibility = merged;
}

/**
 * Waterfall options of the series at `index`: the point indexes drawn as totals (`subtotals`, an
 * empty list removes them) and whether connector lines show. Semantics: PowerPoint.
 */
export function setChartExWaterfall(
	space: ChartExSpace,
	index: number,
	options: { subtotals?: readonly number[]; connectorLines?: boolean },
): ChartExSpace {
	return editLayout(space, index, (layout) => {
		if (options.connectorLines !== undefined)
			setVisibility(layout, { connectorLines: options.connectorLines });
		if (options.subtotals) {
			const subtotals = [...new Set(options.subtotals)]
				.filter((point) => Number.isSafeInteger(point) && point >= 0)
				.sort((a, b) => a - b);
			if (subtotals.length) layout.subtotals = subtotals;
			else delete layout.subtotals;
		}
	});
}

/**
 * Histogram binning of the series at `index`: by category (`'category'`), or numeric bins with
 * either a bin size or a bin count (the size wins, as PowerPoint writes only one). Semantics:
 * PowerPoint.
 */
export function setChartExBinning(
	space: ChartExSpace,
	index: number,
	binning: ChartExBinning | 'category',
): ChartExSpace {
	if (binning !== 'category') {
		if (
			binning.binSize !== undefined &&
			(!Number.isFinite(binning.binSize) || binning.binSize <= 0)
		)
			throw new RangeError('Bin size must be a positive finite number.');
		if (
			binning.binCount !== undefined &&
			(!Number.isSafeInteger(binning.binCount) || binning.binCount <= 0)
		)
			throw new RangeError('Bin count must be a positive integer.');
		for (const limit of [binning.underflow, binning.overflow])
			if (limit !== undefined && limit !== 'auto' && !Number.isFinite(limit))
				throw new RangeError('Bin limits must be finite numbers or auto.');
	}
	return editLayout(space, index, (layout) => {
		if (binning === 'category') {
			delete layout.binning;
			layout.aggregation = true;
			return;
		}
		delete layout.aggregation;
		const next: ChartExBinning = { ...binning };
		if (next.binSize !== undefined) delete next.binCount;
		layout.binning = next;
	});
}

/** Treemap parent label layout (`none`, `banner`, `overlapping`) of the series at `index`. */
export function setChartExParentLabelLayout(
	space: ChartExSpace,
	index: number,
	parentLabelLayout: 'none' | 'banner' | 'overlapping',
): ChartExSpace {
	return editLayout(space, index, (layout) => {
		layout.parentLabelLayout = parentLabelLayout;
	});
}

/** Box and whisker options of the series at `index`. Semantics: PowerPoint. */
export function setChartExBoxWhisker(
	space: ChartExSpace,
	index: number,
	options: {
		quartileMethod?: 'inclusive' | 'exclusive';
		meanLine?: boolean;
		meanMarker?: boolean;
		nonoutliers?: boolean;
		outliers?: boolean;
	},
): ChartExSpace {
	return editLayout(space, index, (layout) => {
		const { quartileMethod, ...flags } = options;
		if (Object.values(flags).some((value) => value !== undefined)) setVisibility(layout, flags);
		if (quartileMethod) layout.statistics = { ...layout.statistics, quartileMethod };
	});
}
