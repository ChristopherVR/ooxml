// Serializable ChartEx requests shared by product controls and headless consumers.
import type { ChartExAxis, ChartExDataLabel, ChartExSpace } from '../ex/model';
import {
	addChartExSeries,
	removeChartExSeries,
	setChartExAxis,
	setChartExLayout,
	setChartExPointLabel,
	type ChartExLayout,
} from './chartex-structure';
import {
	chartExAxisAt,
	editChartExSpace,
	setChartExTitle,
	setChartExWaterfall,
	setChartExBinning,
	setChartExParentLabelLayout,
	setChartExBoxWhisker,
} from './chartex';

export type ChartExEditAction =
	| { op: 'title'; text: string | undefined }
	| { op: 'layout'; layout: ChartExLayout }
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
			return setChartExLayout(space, action.layout);
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
		case 'pointLabel':
			return setChartExPointLabel(space, action.seriesIndex, action.pointIndex, action.label);
	}
}
