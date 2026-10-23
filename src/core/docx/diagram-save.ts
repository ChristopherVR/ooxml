// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Saving edited and inserted SmartArt. An unedited diagram never reaches this module: its run and
// parts are kept byte for byte. An edited diagram keeps its `w:drawing` markup (only the extent can
// change) while its data part is rewritten by the neutral writer and its cached drawing is dropped
// (relationship, part and content type), because Word ignores a stale drawing and lays the diagram
// out again (docs/agnostic-core-plan.md). An inserted diagram gets a new `wp:inline` with
// `dgm:relIds` and four new parts.
import type JSZip from 'jszip';
import {
	DIAGRAM_CONTENT_TYPES,
	DIAGRAM_RELATIONSHIP_TYPES,
	writeEditedDiagramData,
	writeFabricatedDiagramData,
} from '../diagram/index';
import { createOfficeGuid } from '../crypto/uuid';
import type { DocxDiagram } from './diagram';
import { diagramsIn } from './diagram-document';
import type { Block, PendingMediaPart } from './model';
import { buildRelationshipsXml, parseRelationships } from './package-parts';
import { relationshipsPartFor } from './part-relationships';
import type { RelationshipAllocator } from './relationship-allocator';
import { makeNS, makeW, REL_NS, type XmlDocument, type XmlElement } from './xml';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const DGM_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const EMU_PER_PIXEL = 9525;
const CONTENT_TYPES_PART = '[Content_Types].xml';
const HOST_PART = 'word/document.xml';

/** Whether `next` is `base` with edited nodes (same diagram, same data part). */
export function isSameDiagram(base: DocxDiagram | undefined, next: DocxDiagram): boolean {
	return (
		Boolean(base?.parts.data?.partName) &&
		base?.parts.data?.partName === next.parts.data?.partName &&
		!next.inserted
	);
}

const nodesOf = (diagram: DocxDiagram) =>
	diagram.nodes.map(({ id, text, parentId, nodeType }) => ({ id, text, parentId, nodeType }));

/** The data part of an inserted diagram: a document point and the nodes, no presentation points. */
export function insertedDiagramDataXml(
	diagram: DocxDiagram,
	createModelId: () => string = createOfficeGuid,
): string {
	const skeleton = writeFabricatedDiagramData(
		{ nodes: [] },
		{
			layoutUniqueId: diagram.layout?.uniqueId ?? '',
			layoutCategory: diagram.layout?.family ?? 'list',
			quickStyleUniqueId: diagram.quickStyleId ?? '',
			colorsUniqueId: diagram.colorsId ?? '',
			colorsCategory: 'accent1',
		},
		{ createModelId },
	);
	return writeEditedDiagramData(skeleton, nodesOf(diagram), { createModelId });
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

/** Removes a relationship from the document's `.rels` and the part it targets. */
async function dropDrawingPart(zip: JSZip, relId: string, partName: string): Promise<void> {
	const relsPath = relationshipsPartFor(HOST_PART);
	const rels = parseRelationships(await zip.file(relsPath)?.async('string'));
	if (rels.delete(relId)) zip.file(relsPath, buildRelationshipsXml(rels));
	// Another relationship (from any part) may still use the drawing: then keep it.
	for (const rel of rels.values())
		if (rel.target.replace(/^\.?\//, '') === partName.replace(/^word\//, '')) return;
	zip.remove(partName);
	const types = await zip.file(CONTENT_TYPES_PART)?.async('string');
	if (types)
		zip.file(
			CONTENT_TYPES_PART,
			types.replace(
				new RegExp(`<Override\\s+PartName="/${partName.replace(/[.]/g, '\\.')}"[^>]*/>`),
				'',
			),
		);
}

/**
 * Writes the data parts of body diagrams whose nodes or layout changed since `baseBlocks`, and
 * drops their cached drawings. A changed layout also replaces the layout part with the definition
 * supplied in `pendingMedia` under that part's name. Unchanged diagrams are not touched.
 */
export async function applyDiagramEdits(
	zip: JSZip,
	blocks: Block[],
	baseBlocks: Block[],
	pendingMedia?: ReadonlyMap<string, PendingMediaPart>,
): Promise<void> {
	const base = new Map<string, DocxDiagram>();
	for (const diagram of diagramsIn(baseBlocks))
		if (diagram.parts.data?.partName) base.set(diagram.parts.data.partName, diagram);
	const done = new Set<string>();
	for (const diagram of diagramsIn(blocks)) {
		const partName = diagram.parts.data?.partName;
		const original = partName ? base.get(partName) : undefined;
		if (!partName || !original || diagram.inserted || done.has(partName)) continue;
		const layoutId = diagram.layout?.uniqueId;
		const relaid = Boolean(layoutId) && layoutId !== original.layout?.uniqueId;
		if (!relaid && JSON.stringify(nodesOf(diagram)) === JSON.stringify(nodesOf(original))) continue;
		done.add(partName);
		const xml = await zip.file(partName)?.async('string');
		if (xml === undefined) continue;
		const drawing = original.parts.drawing;
		const layoutPart = diagram.parts.layout?.partName;
		const definition = layoutPart ? pendingMedia?.get(layoutPart) : undefined;
		if (relaid && (!layoutPart || !definition))
			throw new Error(
				'The new SmartArt layout definition is missing; the layout cannot be changed.',
			);
		if (relaid && layoutPart && definition) zip.file(layoutPart, definition.bytes);
		zip.file(
			partName,
			writeEditedDiagramData(xml, nodesOf(diagram), {
				dropDrawingLink: true,
				...(relaid && layoutId
					? { layout: { uniqueId: layoutId, category: diagram.layout?.family } }
					: {}),
			}),
		);
		if (drawing) await dropDrawingPart(zip, drawing.relId, drawing.partName);
	}
}

/** Patches the extent of an edited diagram's kept drawing markup when its size changed. */
export function patchDiagramExtent(drawing: XmlElement, widthPx: number, heightPx: number): void {
	const extent = drawing.getElementsByTagNameNS(WP_NS, 'extent')[0];
	if (!extent) return;
	extent.setAttribute('cx', String(Math.max(1, Math.round(widthPx * EMU_PER_PIXEL))));
	extent.setAttribute('cy', String(Math.max(1, Math.round(heightPx * EMU_PER_PIXEL))));
}
