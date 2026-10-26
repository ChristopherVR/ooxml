// ChartEx uses the shared DrawingML style operations through a style-only ChartSpace projection.
// Data, layout properties, axes semantics and label visibility never pass through this projection.
import type { ChartSpace } from '../model';
import type { ChartTextBody, ChartShapeProperties, ChartDataLabels } from '../model-series';
import type { ChartExSpace, ChartExDataLabels } from '../ex/model';
import { editChartExSpace } from './chartex';
import { applyChartAppearance, type ChartAppearance } from './apply-appearance';
import type { ChartFormatTarget } from './target';
type Style = { spPr?: ChartShapeProperties; txPr?: ChartTextBody };
const styles = (holder: Style): Style => ({
	...(holder.spPr ? { spPr: holder.spPr } : {}),
	...(holder.txPr ? { txPr: holder.txPr } : {}),
});
const equal = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);
function labels(source: ChartExDataLabels | undefined): ChartDataLabels | undefined {
	return source
		? {
				...styles(source),
				...(source.position ? { position: source.position } : {}),
				...(source.visibility?.value === undefined ? {} : { showValue: source.visibility.value }),
				...(source.visibility?.categoryName === undefined
					? {}
					: { showCategoryName: source.visibility.categoryName }),
				...(source.visibility?.seriesName === undefined
					? {}
					: { showSeriesName: source.visibility.seriesName }),
				labels: source.labels.map((label) => ({
					...styles(label),
					...(label.index === undefined ? {} : { index: label.index }),
				})),
			}
		: undefined;
}
export function chartExStyleSpace(space: ChartExSpace): ChartSpace {
	return {
		...styles(space),
		...(space.title
			? {
					title: {
						...styles(space.title),
						...(space.title.tx?.rich
							? { tx: { rich: space.title.tx.rich, text: space.title.tx.rich.text } }
							: {}),
					},
				}
			: {}),
		...(space.legend
			? {
					legend: {
						...styles(space.legend),
						entries: [],
						...(space.legend.position ? { position: space.legend.position } : {}),
						...(space.legend.overlay === undefined ? {} : { overlay: space.legend.overlay }),
					},
				}
			: {}),
		plotArea: {
			...styles(space.plotArea),
			axes: space.plotArea.axes.map((axis) => ({
				kind: axis.scaling?.kind ?? 'val',
				scaling: {},
				...styles(axis),
				...(axis.title
					? {
							title: {
								...styles(axis.title),
								...(axis.title.tx?.rich
									? { tx: { rich: axis.title.tx.rich, text: axis.title.tx.rich.text } }
									: {}),
							},
						}
					: {}),
				...(axis.id === undefined ? {} : { id: axis.id }),
				majorGridlines: !!axis.majorGridlines,
				minorGridlines: !!axis.minorGridlines,
				...(axis.majorGridlines?.spPr ? { majorGridlinesSpPr: axis.majorGridlines.spPr } : {}),
				...(axis.minorGridlines?.spPr ? { minorGridlinesSpPr: axis.minorGridlines.spPr } : {}),
			})),
			groups: space.plotArea.series.map((series) => ({
				kind: 'bar',
				element: 'chartex',
				is3D: false,
				axisIds: [],
				...(series.dataLabels ? { dataLabels: labels(series.dataLabels)! } : {}),
				series: [
					{
						...styles(series),
						tx: { text: series.tx?.value ?? series.tx?.rich?.text ?? '' },
						dataPoints: [],
						...(series.dataLabels ? { dataLabels: labels(series.dataLabels)! } : {}),
					},
				],
			})),
		},
	};
}
function copyStyle(target: Style, source: Style): void {
	for (const key of ['spPr', 'txPr'] as const) {
		if (source[key] === undefined) delete target[key];
		else Object.assign(target, { [key]: structuredClone(source[key]) });
	}
}
/** Reuse classic shared style logic only. ChartEx data and semantic model nodes stay untouched. */
export function editChartExStyle(
	space: ChartExSpace,
	operation: (style: ChartSpace) => ChartSpace,
): ChartExSpace {
	const source = chartExStyleSpace(space),
		after = operation(source);
	return editChartExSpace(space, (draft) => {
		copyStyle(draft, after);
		copyStyle(draft.plotArea, after.plotArea);
		if (draft.title && after.title) {
			copyStyle(draft.title, after.title);
			if (after.title.tx?.rich) draft.title.tx = { ...draft.title.tx, rich: after.title.tx.rich };
		}
		if (draft.legend && after.legend) copyStyle(draft.legend, after.legend);
		draft.plotArea.axes.forEach((axis, index) => {
			const edited = after.plotArea.axes[index];
			if (!edited) return;
			copyStyle(axis, edited);
			if (axis.title && edited.title) {
				copyStyle(axis.title, edited.title);
				if (edited.title.tx?.rich) axis.title.tx = { ...axis.title.tx, rich: edited.title.tx.rich };
			}
			for (const key of ['majorGridlines', 'minorGridlines'] as const) {
				const shape =
					edited[key === 'majorGridlines' ? 'majorGridlinesSpPr' : 'minorGridlinesSpPr'];
				if (shape) {
					axis[key] ??= {};
					axis[key]!.spPr = structuredClone(shape);
				}
			}
		});
		draft.plotArea.series.forEach((series, index) => {
			const group = after.plotArea.groups[index],
				old = source.plotArea.groups[index];
			const edited = group?.series[0];
			if (!edited || !old) return;
			copyStyle(series, edited);
			const chosen = !equal(group.dataLabels, old.dataLabels)
				? group.dataLabels
				: edited.dataLabels;
			if (!chosen) return;
			series.dataLabels ??= { labels: [], hidden: [] };
			copyStyle(series.dataLabels, chosen);
			for (const label of chosen.labels) {
				const target = series.dataLabels.labels.find((item) => item.index === label.index);
				if (target) copyStyle(target, label);
			}
		});
	});
}
export function applyChartExAppearance(
	space: ChartExSpace,
	target: ChartFormatTarget,
	before: ChartAppearance | undefined,
	after: ChartAppearance,
): ChartExSpace {
	return editChartExStyle(space, (snapshot) =>
		applyChartAppearance(snapshot, target, before, after),
	);
}
