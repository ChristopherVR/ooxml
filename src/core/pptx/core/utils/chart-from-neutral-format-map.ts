import type { ChartShapeProperties, ChartSpace } from '../../../chart/index';
import type { XmlObject } from '../types';
import { pairGroupNodes, rawChildren } from './chart-from-neutral';
import type { ChartTreeReader } from './chart-series-tree-fields';

/** Associate compatibility source nodes with their authoritative neutral primitive models. */
export function neutralChartShapeMap(
	space: ChartSpace,
	source: XmlObject | undefined,
	reader: ChartTreeReader,
): (source: XmlObject) => ChartShapeProperties | undefined {
	const shapes = new WeakMap<XmlObject, ChartShapeProperties>();
	const child = (node: XmlObject | undefined, name: string) =>
		reader.xml.getChildByLocalName(node, name);
	const put = (owner: XmlObject | undefined, shape: ChartShapeProperties | undefined) => {
		const sourceShape = child(owner, 'spPr');
		if (sourceShape) shapes.set(sourceShape, shape ?? {});
	};
	const chart = child(source, 'chart');
	const plot = child(chart, 'plotArea');
	put(source, space.spPr);
	put(plot, space.plotArea.spPr);
	put(child(chart, 'title'), space.title?.spPr);
	put(child(chart, 'legend'), space.legend?.spPr);
	for (const axis of space.plotArea.axes) {
		const node = rawChildren(plot, `${axis.kind}Ax`, reader.localName).find(
			(candidate) => Number(child(candidate, 'axId')?.['@_val']) === axis.id,
		);
		put(node, axis.spPr);
		put(child(node, 'majorGridlines'), axis.majorGridlinesSpPr);
		put(child(node, 'minorGridlines'), axis.minorGridlinesSpPr);
	}
	for (const { group, node } of pairGroupNodes(
		space.plotArea.groups,
		plot ?? {},
		reader.localName,
	)) {
		put(child(node, 'dLbls'), group.dataLabels?.spPr);
		const seriesNodes = rawChildren(node, 'ser', reader.localName);
		group.series.forEach((series, index) => {
			put(seriesNodes[index], series.spPr);
			put(child(seriesNodes[index], 'dLbls'), series.dataLabels?.spPr);
		});
	}
	return (node) => shapes.get(node);
}
