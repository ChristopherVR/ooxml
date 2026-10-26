// The existing format view reads a style-only classic snapshot; all writes target ChartEx itself.
import type {
	ChartAxis,
	ChartDataLabels,
	ChartExDataLabels,
	ChartExSpace,
	ChartShapeProperties,
	ChartSpace,
} from 'ooxml-core/chart';
import type { DrawingColorTheme } from 'ooxml-core/drawingml';
import { chartSpaceFormatSource } from './format-pane-space';
import type { ChartFormatSource } from './format-pane-view';

const paint = (shape: ChartShapeProperties | undefined) => (shape ? { spPr: shape } : {});
function labels(source: ChartExDataLabels): ChartDataLabels {
	return {
		...paint(source.spPr),
		labels: [],
		...(source.position ? { position: source.position } : {}),
		...(source.visibility?.value === undefined ? {} : { showValue: source.visibility.value }),
		...(source.visibility?.categoryName === undefined
			? {}
			: { showCategoryName: source.visibility.categoryName }),
		...(source.visibility?.seriesName === undefined
			? {}
			: { showSeriesName: source.visibility.seriesName }),
	};
}
/** Shared formatting controls for ChartEx fills, outlines, legend and label content. */
export function chartExSpaceFormatSource(
	space: ChartExSpace,
	theme?: DrawingColorTheme,
): ChartFormatSource {
	const axes: ChartAxis[] = space.plotArea.axes.map((axis) => ({
		kind: axis.scaling?.kind ?? 'val',
		scaling: {},
		majorGridlines: !!axis.majorGridlines,
		minorGridlines: !!axis.minorGridlines,
		...paint(axis.spPr),
		...(axis.id === undefined ? {} : { id: axis.id }),
		...(axis.majorGridlines?.spPr ? { majorGridlinesSpPr: axis.majorGridlines.spPr } : {}),
		...(axis.minorGridlines?.spPr ? { minorGridlinesSpPr: axis.minorGridlines.spPr } : {}),
	}));
	const snapshot: ChartSpace = {
		...paint(space.spPr),
		plotArea: {
			...paint(space.plotArea.spPr),
			axes,
			groups: [
				{
					kind: 'bar',
					element: 'chartex',
					is3D: false,
					axisIds: [],
					series: space.plotArea.series.map((series, index) => ({
						...paint(series.spPr),
						dataPoints: [],
						tx: { text: series.tx?.value ?? series.tx?.rich?.text ?? `Series ${index + 1}` },
						...(series.dataLabels ? { dataLabels: labels(series.dataLabels) } : {}),
					})),
				},
			],
		},
		...(space.title
			? {
					title: {
						...paint(space.title.spPr),
						text: space.title.tx?.value ?? space.title.tx?.rich?.text ?? '',
					},
				}
			: {}),
		...(space.legend
			? {
					legend: {
						...paint(space.legend.spPr),
						entries: [],
						...(space.legend.position ? { position: space.legend.position } : {}),
						...(space.legend.overlay === undefined ? {} : { overlay: space.legend.overlay }),
					},
				}
			: {}),
	};
	const options = { ...(theme ? { theme } : {}), advanced: false };
	const source = chartSpaceFormatSource(snapshot, options);
	// The classic axis/series editors expose fields with different ChartEx semantics.
	const { axisOptions: _axes, seriesOptions: _series, ...common } = source;
	return common;
}
