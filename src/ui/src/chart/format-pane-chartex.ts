// ChartEx shares DrawingML formatting; its family and axis controls keep native semantics.
import { chartExStyleSpace, type ChartExSpace } from 'ooxml-core/chart';
import type { DrawingColorTheme } from 'ooxml-core/drawingml';
import { chartSpaceFormatSource } from './format-pane-space';
import type { ChartFormatSource } from './format-pane-view';
export function chartExSpaceFormatSource(
	space: ChartExSpace,
	theme?: DrawingColorTheme,
): ChartFormatSource {
	const source = chartSpaceFormatSource(chartExStyleSpace(space), {
		...(theme ? { theme } : {}),
		advanced: true,
	});
	const { axisOptions: _axes, seriesOptions: _series, ...common } = source;
	return {
		...common,
		element(target) {
			const view = source.element(target);
			return { ...view, support: { ...view.support, secondaryAxis: false } };
		},
	};
}
