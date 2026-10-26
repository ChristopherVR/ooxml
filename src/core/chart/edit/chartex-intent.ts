// Map the shared Format pane's supported requests onto ChartEx without rewriting its data model.
import type { ChartExSpace } from '../ex/model';
import type { ChartShapeProperties } from '../model-series';
import type { ChartEditIntent } from './intent';
import type { ChartFormatTarget } from './target';
import { chartExSeriesAt, editChartExSpace, setChartExLegend } from './chartex';
import { withDrawingColorTransparency } from './paint';

function shapes(space: ChartExSpace, target: ChartFormatTarget): ChartShapeProperties[] {
	switch (target.kind) {
		case 'chartArea':
			return [(space.spPr ??= {})];
		case 'plotArea':
			return [(space.plotArea.spPr ??= {})];
		case 'title':
			if (!space.title) throw new Error('The chart has no title.');
			return [(space.title.spPr ??= {})];
		case 'legend':
			if (!space.legend) throw new Error('The chart has no legend.');
			return [(space.legend.spPr ??= {})];
		case 'series':
			return [(chartExSeriesAt(space, target.seriesIndex).spPr ??= {})];
		case 'axis':
		case 'majorGridlines':
		case 'minorGridlines': {
			const axis = space.plotArea.axes[target.axisIndex];
			if (!axis) throw new RangeError('No chart axis at index.');
			if (target.kind === 'axis') return [(axis.spPr ??= {})];
			return [((axis[target.kind] ??= {}).spPr ??= {})];
		}
		case 'dataLabels': {
			const series =
				target.seriesIndex === undefined
					? space.plotArea.series
					: [chartExSeriesAt(space, target.seriesIndex)];
			return series.map((series) => {
				series.dataLabels ??= { labels: [], hidden: [] };
				return (series.dataLabels.spPr ??= {});
			});
		}
	}
}

export function applyChartExEditIntent(space: ChartExSpace, intent: ChartEditIntent): ChartExSpace {
	if (intent.op === 'legend') {
		let next = setChartExLegend(space, {
			visible: intent.edit.visible ?? true,
			...(intent.edit.position ? { position: intent.edit.position } : {}),
		});
		const overlay = intent.edit.overlay;
		if (overlay !== undefined && next.legend)
			next = editChartExSpace(next, (d) => {
				d.legend!.overlay = overlay;
			});
		return next;
	}
	if (intent.op === 'fill' || intent.op === 'line' || intent.op === 'transparency')
		return editChartExSpace(space, (draft) => {
			for (const shape of shapes(draft, intent.target)) {
				if (intent.op === 'fill') {
					if (intent.fill) shape.fill = structuredClone(intent.fill);
					else delete shape.fill;
				} else if (intent.op === 'line') {
					if (intent.line)
						shape.line = intent.merge
							? { ...shape.line, ...structuredClone(intent.line) }
							: structuredClone(intent.line);
					else delete shape.line;
				} else {
					const holder = intent.paint === 'fill' ? shape : shape.line;
					if (holder?.fill?.kind === 'solid')
						holder.fill = {
							kind: 'solid',
							color: withDrawingColorTransparency(holder.fill.color, intent.percent),
						};
				}
			}
		});
	if (intent.op === 'dataLabels')
		return editChartExSpace(space, (draft) => {
			const edit = intent.edit;
			if (edit.showPercent || edit.showBubbleSize || edit.showLegendKey)
				throw new Error('The extended chart does not support this label content.');
			const selected =
				intent.seriesIndex === undefined
					? draft.plotArea.series
					: [chartExSeriesAt(draft, intent.seriesIndex)];
			for (const series of selected) {
				if (edit.show === false) {
					delete series.dataLabels;
					continue;
				}
				const labels = (series.dataLabels ??= { labels: [], hidden: [] });
				labels.visibility ??= {};
				if (edit.showValue !== undefined) labels.visibility.value = edit.showValue;
				if (edit.showCategoryName !== undefined)
					labels.visibility.categoryName = edit.showCategoryName;
				if (edit.showSeriesName !== undefined) labels.visibility.seriesName = edit.showSeriesName;
				if (edit.position !== undefined) {
					if (edit.position) labels.position = edit.position;
					else delete labels.position;
				}
				if (edit.separator !== undefined) {
					if (edit.separator !== null) labels.separator = edit.separator;
					else delete labels.separator;
				}
				if (edit.numberFormat !== undefined) {
					if (edit.numberFormat)
						labels.numberFormat = { formatCode: edit.numberFormat, sourceLinked: false };
					else delete labels.numberFormat;
				}
			}
		});
	throw new Error(`The extended chart does not support the ${intent.op} edit.`);
}
