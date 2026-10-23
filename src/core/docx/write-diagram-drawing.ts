// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// The run of a SmartArt graphic: an edited diagram keeps its `w:drawing` markup (only the extent
// can change; its parts are rewritten by diagram-save.ts) and an inserted one gets a new
// `wp:inline` with `dgm:relIds`, registering its four parts with the relationship allocator.
import { DIAGRAM_CONTENT_TYPES, DIAGRAM_RELATIONSHIP_TYPES } from '../diagram/index';
import type { DocxDiagram } from './diagram';
import { insertedDiagramDataXml } from './diagram-save';
import type { InlineImage } from './model';
import type { RelationshipAllocator } from './relationship-allocator';
import { clearNonProperties } from './write-drawing-patch';
import { makeNS, makeW, REL_NS, type XmlDocument, type XmlElement } from './xml';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const DGM_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const EMU_PER_PIXEL = 9525;

/** Whether `next` is `base` with edited nodes (same diagram, same data part). */
export function isSameDiagram(base: DocxDiagram | undefined, next: DocxDiagram): boolean {
	return (
		Boolean(base?.parts.data?.partName) &&
		base?.parts.data?.partName === next.parts.data?.partName &&
		!next.inserted
	);
}

function el(doc: XmlDocument, ns: string, qualified: string, attrs: Record<string, string> = {}) {
	const node = makeNS(doc, ns, qualified);
	for (const [name, value] of Object.entries(attrs)) node.setAttribute(name, value);
	return node;
}

/** Registers the parts of an inserted diagram and builds its `w:drawing` (`wp:inline`). */
export function buildInsertedDiagramDrawing(
	doc: XmlDocument,
	diagram: DocxDiagram,
	widthPx: number,
	heightPx: number,
	allocator: RelationshipAllocator,
): XmlElement {
	const part = (name: 'data' | 'layout' | 'quickStyle' | 'colors', content?: string): string => {
		const partName = diagram.parts[name]?.partName;
		if (!partName) throw new Error(`The inserted SmartArt has no ${name} part.`);
		return allocator.addDiagramPart({
			partName,
			type: DIAGRAM_RELATIONSHIP_TYPES[name],
			contentType: DIAGRAM_CONTENT_TYPES[name],
			...(content !== undefined ? { content } : {}),
		});
	};
	const dm = part('data', insertedDiagramDataXml(diagram));
	const lo = part('layout');
	const qs = part('quickStyle');
	const cs = part('colors');
	const id = allocator.docPrIds.next();
	const cx = String(Math.max(1, Math.round(widthPx * EMU_PER_PIXEL)));
	const cy = String(Math.max(1, Math.round(heightPx * EMU_PER_PIXEL)));
	const drawing = makeW(doc, 'drawing');
	const inline = el(doc, WP_NS, 'wp:inline', { distT: '0', distB: '0', distL: '0', distR: '0' });
	inline.appendChild(el(doc, WP_NS, 'wp:extent', { cx, cy }));
	inline.appendChild(el(doc, WP_NS, 'wp:effectExtent', { l: '0', t: '0', r: '0', b: '0' }));
	inline.appendChild(
		el(doc, WP_NS, 'wp:docPr', { id, name: `${diagram.name ?? 'Diagram'} ${id}` }),
	);
	inline.appendChild(el(doc, WP_NS, 'wp:cNvGraphicFramePr'));
	const graphic = el(doc, A_NS, 'a:graphic');
	const graphicData = el(doc, A_NS, 'a:graphicData', { uri: DGM_NS });
	const relIds = el(doc, DGM_NS, 'dgm:relIds');
	relIds.setAttributeNS(REL_NS, 'r:dm', dm);
	relIds.setAttributeNS(REL_NS, 'r:lo', lo);
	relIds.setAttributeNS(REL_NS, 'r:qs', qs);
	relIds.setAttributeNS(REL_NS, 'r:cs', cs);
	graphicData.appendChild(relIds);
	graphic.appendChild(graphicData);
	inline.appendChild(graphic);
	drawing.appendChild(inline);
	return drawing;
}

/** Patches the extent of an edited diagram's kept drawing markup when its size changed. */
export function patchDiagramExtent(drawing: XmlElement, widthPx: number, heightPx: number): void {
	const extent = drawing.getElementsByTagNameNS(WP_NS, 'extent')[0];
	if (!extent) return;
	extent.setAttribute('cx', String(Math.max(1, Math.round(widthPx * EMU_PER_PIXEL))));
	extent.setAttribute('cy', String(Math.max(1, Math.round(heightPx * EMU_PER_PIXEL))));
}

/** The run for a SmartArt image, or undefined when `image` is not an edited or inserted one. */
export function diagramRun(
	doc: XmlDocument,
	node: XmlElement,
	image: InlineImage,
	base: InlineImage | undefined,
	oldDrawing: XmlElement | undefined,
	oldContainer: XmlElement | undefined,
	allocator: RelationshipAllocator | undefined,
): XmlElement | undefined {
	if (!image.diagram) return undefined;
	if (base?.diagram && oldDrawing && isSameDiagram(base.diagram, image.diagram)) {
		if (image.widthPx !== base.widthPx || image.heightPx !== base.heightPx)
			patchDiagramExtent(oldDrawing, image.widthPx, image.heightPx);
		clearNonProperties(node, oldContainer);
		return node;
	}
	if (!base && image.diagram.inserted) {
		if (!allocator) throw new Error('Cannot insert SmartArt without a relationship allocator.');
		clearNonProperties(node);
		node.appendChild(
			buildInsertedDiagramDrawing(doc, image.diagram, image.widthPx, image.heightPx, allocator),
		);
		return node;
	}
	return undefined;
}
