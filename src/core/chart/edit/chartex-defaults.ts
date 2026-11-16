// The look Office gives a new extended chart of each family, from the parts Excel 16 writes for
// Insert > Chart (`Shapes.AddChart2(-1, type)`, observed output): which data labels show, the
// axis gaps and gridlines, the legend and the family options. Word and Excel apply it to the
// charts they insert, so a new chart opens in Office looking as Office would have drawn it.
import type { ChartExSpace } from '../ex/model';
import { editChartExSpace } from './chartex';
import type { ChartExFamily } from './chartex-family';

const labels = (position: string | undefined, show: 'categoryName' | 'value') => ({
	...(position ? { position } : {}),
	visibility: { seriesName: false, categoryName: show === 'categoryName', value: show === 'value' },
	labels: [],
	hidden: [],
});

/** Applies Office's new-chart defaults of `family` to a chart made by `createChartExSpace`. */
export function applyChartExOfficeDefaults(
	space: ChartExSpace,
	family: ChartExFamily,
): ChartExSpace {
	return editChartExSpace(space, (draft) => {
		const series = draft.plotArea.series.filter((s) => s.ownerIndex === undefined);
		const cat = draft.plotArea.axes.find((axis) => axis.scaling?.kind === 'cat');
		const val = draft.plotArea.axes.find(
			(axis) => axis.scaling?.kind === 'val' && axis.units?.unit !== 'percentage',
		);
		const gap = (gapWidth: number) => {
			if (cat) cat.scaling = { kind: 'cat', gapWidth };
		};
		if (val) val.majorGridlines ??= {};
		const legend = () => {
			draft.legend ??= { position: 't', align: 'ctr', overlay: false };
		};
		for (const s of series) {
			const layout = (s.layoutProperties ??= {});
			switch (family) {
				case 'waterfall':
					s.dataLabels ??= labels('outEnd', 'value');
					layout.subtotals ??= [];
					break;
				case 'clusteredColumn':
					if (layout.binning) layout.binning.intervalClosed ??= 'r';
					break;
				case 'boxWhisker':
					layout.visibility = {
						meanLine: false,
						meanMarker: true,
						nonoutliers: false,
						outliers: true,
						...layout.visibility,
					};
					layout.statistics ??= { quartileMethod: 'exclusive' };
					break;
				case 'treemap':
					s.dataLabels ??= labels('inEnd', 'categoryName');
					layout.parentLabelLayout ??= 'overlapping';
					break;
				case 'sunburst':
					s.dataLabels ??= labels('ctr', 'categoryName');
					break;
				case 'funnel':
					s.dataLabels ??= labels(undefined, 'value');
					break;
			}
			if (!Object.keys(layout).length) delete s.layoutProperties;
		}
		if (family === 'waterfall') gap(0.5);
		if (family === 'clusteredColumn' || family === 'pareto') gap(0);
		if (family === 'boxWhisker') gap(1);
		if (family === 'funnel' && !draft.plotArea.axes.length) {
			draft.plotArea.axes = [{ id: 1, scaling: { kind: 'cat', gapWidth: 0.06 }, tickLabels: {} }];
			for (const s of series) s.axisIds = [];
		}
		if (family === 'waterfall' || family === 'treemap') legend();
	});
}
