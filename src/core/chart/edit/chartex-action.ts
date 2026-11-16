// Serializable ChartEx requests shared by product controls and headless consumers.
import type { ChartExAxis, ChartExDataLabel, ChartExSpace } from '../ex/model';
import {
	addChartExSeries,
	removeChartExSeries,
	setChartExAxis,
	setChartExPointLabel,
} from './chartex-structure';
import { setChartExFamily, type ChartExFamily } from './chartex-family';
import { editChartExAxis, type ChartExAxisEdit } from './chartex-axis';
import { setChartExRegionMap, type ChartExRegionMapOptions } from './chartex-region-map';
import { chartExDataSeriesIndices } from './chartex-sheet';
import {
	chartExAxisAt,
	editChartExSpace,
	setChartExTitle,
	setChartExWaterfall,
	setChartExBinning,
	setChartExParentLabelLayout,
	setChartExBoxWhisker,
	renameChartExSeries,
} from './chartex';

export type ChartExEditAction =
	| { op: 'title'; text: string | undefined }
	| { op: 'layout'; layout: ChartExFamily }
	/** Renames the data series at `seriesIndex` (a plot position); a linked name keeps its cell. */
	| { op: 'seriesName'; seriesIndex: number; name: string }
	| { op: 'addSeries'; name: string; templateIndex?: number }
	| { op: 'removeSeries'; seriesIndex: number }
	| { op: 'waterfall'; seriesIndex: number; options: Parameters<typeof setChartExWaterfall>[2] }
	| { op: 'binning'; seriesIndex: number; binning: Parameters<typeof setChartExBinning>[2] }
	| {
			op: 'parentLabels';
			seriesIndex: number;
			layout: Parameters<typeof setChartExParentLabelLayout>[2];
	  }
	| { op: 'boxWhisker'; seriesIndex: number; options: Parameters<typeof setChartExBoxWhisker>[2] }
	| { op: 'axis'; axisIndex: number; edit: Partial<ChartExAxis> }
	| { op: 'tickLabels'; axisIndex: number; visible: boolean }
	/** Axis title, visibility, scaling, units, number format, gridlines and tick marks. */
	| { op: 'axisOptions'; axisIndex: number; edit: ChartExAxisEdit }
	| { op: 'regionMap'; seriesIndex: number; options: ChartExRegionMapOptions }
	| {
			op: 'pointLabel';
			seriesIndex: number;
			pointIndex: number;
			label: Omit<ChartExDataLabel, 'index'> | undefined;
	  };

export function applyChartExEditAction(
	space: ChartExSpace,
	action: ChartExEditAction,
): ChartExSpace {
	switch (action.op) {
		case 'title':
			return setChartExTitle(space, action.text);
		case 'layout':
			return setChartExFamily(space, action.layout);
		case 'seriesName':
			if (!chartExDataSeriesIndices(space).includes(action.seriesIndex))
				throw new RangeError('Only a data series has a name to edit.');
			return renameChartExSeries(space, action.seriesIndex, action.name);
		case 'addSeries':
			return addChartExSeries(space, action.name, action.templateIndex);
		case 'removeSeries':
			return removeChartExSeries(space, action.seriesIndex);
		case 'waterfall':
			return setChartExWaterfall(space, action.seriesIndex, action.options);
		case 'binning':
			return setChartExBinning(space, action.seriesIndex, action.binning);
		case 'parentLabels':
			return setChartExParentLabelLayout(space, action.seriesIndex, action.layout);
		case 'boxWhisker':
			return setChartExBoxWhisker(space, action.seriesIndex, action.options);
		case 'axis':
			return setChartExAxis(space, action.axisIndex, action.edit);
		case 'tickLabels':
			return editChartExSpace(space, (draft) => {
				const axis = chartExAxisAt(draft, action.axisIndex);
				if (action.visible) axis.tickLabels ??= {};
				else delete axis.tickLabels;
			});
		case 'axisOptions':
			return editChartExAxis(space, action.axisIndex, action.edit);
		case 'regionMap':
			return setChartExRegionMap(space, action.seriesIndex, action.options);
		case 'pointLabel':
			return setChartExPointLabel(space, action.seriesIndex, action.pointIndex, action.label);
	}
}
