import { NS, children, first, type XmlElement } from '../xml';

/** Shared element mapping for direct chart formatting, including scatter X/Y axes. */
export function chartFormattingNodes(root: XmlElement) {
	const chart = first(root, 'chart', NS.c);
	const plot = first(chart, 'plotArea', NS.c);
	const category = first(plot, 'catAx', NS.c) ?? first(plot, 'dateAx', NS.c);
	const values = plot ? children(plot, 'valAx', NS.c) : [];
	return {
		chartArea: root,
		plotArea: plot,
		title: first(chart, 'title', NS.c),
		legend: first(chart, 'legend', NS.c),
		categoryAxis: category ?? values[0],
		valueAxis: category ? values[0] : values[1],
		gridlineMajor: first(values[0], 'majorGridlines', NS.c),
		gridlineMinor: first(values[0], 'minorGridlines', NS.c),
	};
}
