import type { ChartDataSheetEdit } from 'ooxml-core/chart';
import { chartDataApplySheetEdits, chartDataToSheet } from 'ooxml-core/pptx';
import type { PptxChartData } from 'ooxml-core/pptx';
import { defineChartDataSheet } from '../../chart/data-sheet';
import type { OfficeChartDataSheetLabels, OfficeUiChartDataSheet } from '../../chart/data-sheet';
import { base, bridged } from './office-aliases';

type Translate = (key: string, params?: Record<string, string | number>) => string;

/** What a binding passes to `pptx-ui-chart-data-sheet` through `state`. */
export interface ChartDataSheetViewState {
	chartData: PptxChartData | undefined;
	canEdit: boolean;
	/**
	 * File > Options > Advanced > "Properties follow chart data point for current workbook":
	 * whether removing a category re-indexes per-point formatting (default `true`).
	 */
	followDataPoint?: boolean;
	/** The chart part picked on the canvas: a value cell, or a series' name cell. */
	highlight?: { seriesIndex: number; pointIndex?: number } | null;
	/** BCP-47 language the values are read and shown in (default: the browser's). */
	locale?: string;
	translate?: Translate;
}

/** Detail of `pptx-chart-data-change`: the chart data after the edits, ready to commit. */
export interface ChartDataSheetChange {
	chartData: PptxChartData;
	edits: ChartDataSheetEdit[];
}
export type ChartDataSheetChangeEvent = CustomEvent<ChartDataSheetChange>;

export interface PptxUiChartDataSheetElement extends OfficeUiChartDataSheet {
	state: ChartDataSheetViewState;
}
declare global {
	interface HTMLElementTagNameMap {
		'pptx-ui-chart-data-sheet': PptxUiChartDataSheetElement;
	}
}

/** The sheet's text from the pptx dictionary (the keys every binding's grid used). */
export function chartDataSheetLabels(t: Translate): OfficeChartDataSheetLabels {
	return {
		heading: t('pptx.chart.data'),
		addCategory: t('pptx.chart.addCategory'),
		addSeries: t('pptx.chart.addSeries'),
		addCategoryShort: t('pptx.chart.cat'),
		addSeriesShort: t('pptx.chart.seriesShort'),
		removeCategory: t('pptx.chart.removeCategory'),
		removeSeries: t('pptx.chart.removeSeries'),
		categories: t('pptx.chart.categories'),
		seriesName: `${t('pptx.chart.seriesShort')} {number}`,
		categoryLabel: `${t('pptx.chart.categories')} {number}`,
		valueCell: t('pptx.chart.seriesValueCell', { name: '{name}', number: '{number}' }),
		summary: `{series} ${t('pptx.chart.series')} · {categories} ${t('pptx.chart.categories')}`,
		invalidNumber: t('pptx.chart.dataSheetInvalidNumber'),
	};
}

const identity: Translate = (key) => key;

/**
 * `pptx-ui-chart-data-sheet`: the shared chart data sheet fed a pptx chart. A binding sets
 * `state` and commits the `pptx-chart-data-change` detail's `chartData` through its own update
 * path; the edits go through the pptx chart grid policy (`chartDataApplySheetEdits`), so series
 * and point formatting survive every edit.
 */
export function definePptxChartDataSheet(registry: CustomElementRegistry): void {
	if (registry.get('pptx-ui-chart-data-sheet')) {
		return;
	}
	const Base = bridged(
		base(registry, defineChartDataSheet, 'office-ui-chart-data-sheet'),
		'pptx-ui-chart-data-sheet',
	) as unknown as new () => OfficeUiChartDataSheet;
	class PptxChartDataSheet extends Base {
		#model: ChartDataSheetViewState = { chartData: undefined, canEdit: false };
		get state(): ChartDataSheetViewState {
			return this.#model;
		}
		set state(value: ChartDataSheetViewState) {
			this.#model = value;
			const chartData = value.chartData;
			this.hidden = !chartData;
			this.labels = chartDataSheetLabels(value.translate ?? identity);
			this.locale = value.locale ?? '';
			this.disabled = !value.canEdit;
			const mark = value.highlight;
			this.highlight = mark
				? {
						series: mark.seriesIndex,
						...(mark.pointIndex !== undefined ? { point: mark.pointIndex } : {}),
					}
				: null;
			this.data = chartData ? chartDataToSheet(chartData) : { categories: [], series: [] };
		}
		protected override emitEdits(edits: ChartDataSheetEdit[]): void {
			const current = this.#model.chartData;
			const next =
				current && !this.disabled
					? chartDataApplySheetEdits(current, edits, this.#model.followDataPoint ?? true)
					: null;
			if (!next) {
				// Refused: show the chart's data again rather than the edit.
				this.requestUpdate();
				return;
			}
			super.emitEdits(edits);
			this.#model = { ...this.#model, chartData: next };
			const detail: ChartDataSheetChange = { chartData: next, edits };
			this.fire('pptx-chart-data-change', detail);
		}
	}
	registry.define(
		'pptx-ui-chart-data-sheet',
		PptxChartDataSheet as unknown as CustomElementConstructor,
	);
}
