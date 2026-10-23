// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// The inline drawing of a chart inserted in this session: `wp:inline` around an `a:graphicData`
// whose `c:chart` names the chart part's relationship, the markup Word writes for Insert > Chart.
import { EMU_PER_PIXEL } from './drawing';
import { CHART_GRAPHIC_URI } from './chart';
import type { InlineImage } from './model';
import type { RelationshipAllocator } from './relationship-allocator';
import { clearNonProperties, updateExtent } from './write-drawing-patch';
import { makeNS, makeW, REL_NS, type XmlDocument, type XmlElement } from './xml';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

function child(parent: XmlElement, ns: string, name: string, attrs: Record<string, string> = {}) {
	const node = makeNS(parent.ownerDocument!, ns, name);
	for (const [key, value] of Object.entries(attrs)) node.setAttribute(key, value);
	parent.appendChild(node);
	return node;
}

/** `w:drawing` holding an inline chart frame for relationship `relId`. */
export function buildChartDrawing(
	doc: XmlDocument,
	image: InlineImage,
	relId: string,
	docPrId: string,
): XmlElement {
	const drawing = makeW(doc, 'drawing');
	const inline = child(drawing, WP_NS, 'wp:inline', {
		distT: '0',
		distB: '0',
		distL: '0',
		distR: '0',
	});
	child(inline, WP_NS, 'wp:extent', {
		cx: String(Math.max(1, Math.round(image.widthPx * EMU_PER_PIXEL))),
		cy: String(Math.max(1, Math.round(image.heightPx * EMU_PER_PIXEL))),
	});
	child(inline, WP_NS, 'wp:effectExtent', { l: '0', t: '0', r: '0', b: '0' });
	const docPr = child(inline, WP_NS, 'wp:docPr', { id: docPrId, name: `Chart ${docPrId}` });
	if (image.altText) docPr.setAttribute('descr', image.altText);
	if (image.title) docPr.setAttribute('title', image.title);
	child(inline, WP_NS, 'wp:cNvGraphicFramePr');
	const graphic = child(inline, A_NS, 'a:graphic');
	const data = child(graphic, A_NS, 'a:graphicData', { uri: CHART_GRAPHIC_URI });
	const chart = child(data, C_NS, 'c:chart');
	chart.setAttributeNS(REL_NS, 'r:id', relId);
	return drawing;
}

/**
 * The run of a chart, or undefined when `image` is not a chart this module writes. An edited chart
 * keeps its drawing (the part is saved by chart-save.ts) and only its size can change; a chart
 * inserted in this session gets a new drawing and a chart relationship.
 */
export function chartRun(
	doc: XmlDocument,
	node: XmlElement,
	image: InlineImage,
	base: InlineImage | undefined,
	oldDrawing: XmlElement | undefined,
	oldContainer: XmlElement | undefined,
	allocator: RelationshipAllocator | undefined,
): XmlElement | undefined {
	if (
		base?.chart &&
		oldDrawing &&
		image.chart?.partName &&
		image.chart.partName === base.chart.partName &&
		image.chart.relId === base.chart.relId
	) {
		if (image.widthPx !== base.widthPx || image.heightPx !== base.heightPx)
			updateExtent(oldDrawing, image.widthPx, image.heightPx);
		clearNonProperties(node, oldContainer);
		return node;
	}
	if (!base && image.chart?.chartSpace && !image.chart.relId && !image.chart.partName) {
		if (!allocator) throw new Error('Cannot insert a chart without a relationship allocator.');
		const relId = allocator.addChart(image.chart.chartSpace);
		clearNonProperties(node);
		node.appendChild(buildChartDrawing(doc, image, relId, allocator.docPrIds.next()));
		return node;
	}
	return undefined;
}
