import { NS, children, elements, first, type XmlElement } from '../xml';

/**
 * The `c:dLbls` whose text formatting the painter uses for data labels: the first series' (else the
 * first group's) that switches a label on or carries text properties. Read-only: the formatting
 * writers do not patch it.
 */
export function chartDataLabelNode(root: XmlElement): XmlElement | undefined {
	const plot = first(first(root, 'chart', NS.c), 'plotArea', NS.c);
	const groups = plot ? elements(plot).filter((node) => node.localName.endsWith('Chart')) : [];
	const candidates = groups.flatMap((group) => [
		...children(group, 'ser', NS.c).map((ser) => first(ser, 'dLbls', NS.c)),
		first(group, 'dLbls', NS.c),
	]);
	return candidates.find(
		(node) =>
			node !== undefined &&
			(first(node, 'txPr', NS.c) !== undefined ||
				elements(node).some(
					(child) => child.localName.startsWith('show') && child.getAttribute('val') === '1',
				)),
	);
}

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
