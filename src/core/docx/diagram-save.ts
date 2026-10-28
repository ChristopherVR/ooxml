// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Saving edited and inserted SmartArt. An unedited diagram never reaches this module: its run and
// parts are kept byte for byte. An edited diagram keeps its `w:drawing` markup (only the extent can
// change) while its data part is rewritten by the neutral writer and its cached drawing is dropped
// (relationship, part and content type), because Word ignores a stale drawing and lays the diagram
// out again (docs/agnostic-core-plan.md). The run markup is written by write-diagram-drawing.ts.
import type JSZip from 'jszip';
import {
	writeEditedDiagramData,
	writeInsertedDiagramData,
	withDiagramPictureSources,
	diagramPictureRelationships,
} from '../diagram/index';
import { createOfficeGuid } from '../crypto/uuid';
import { removeContentTypeOverride, resolvePartPath } from '../opc/index';
import type { DocxDiagram } from './diagram';
import { diagramsIn } from './diagram-document';
import type { Block, PendingMediaPart } from './model';
import { parseRelationships } from './package-parts';
import { NS, buildXml, elements, parseXml } from '../xml/index';
import { DIAGRAM_RELATIONSHIP_TYPES } from '../diagram/relationships';
import { relationshipsPartFor } from './part-relationships';

const HOST_PART = 'word/document.xml';

const nodesOf = (diagram: DocxDiagram) =>
	diagram.nodes.map(({ id, text, parentId, nodeType }) => ({ id, text, parentId, nodeType }));

/** The data part of an inserted diagram (the shared `writeInsertedDiagramData`). */
export function insertedDiagramDataXml(
	diagram: DocxDiagram,
	createModelId: () => string = createOfficeGuid,
): string {
	return writeInsertedDiagramData(
		{
			nodes: diagram.nodes,
			layoutId: diagram.layout?.uniqueId ?? '',
			layoutCategory: diagram.layout?.family,
			quickStyleId: diagram.quickStyleId ?? '',
			colorsId: diagram.colorsId ?? '',
		},
		createModelId,
	);
}

/** Removes a relationship from the document's `.rels` and the part it targets. */
async function dropDrawingPart(
	zip: JSZip,
	relId: string,
	partName: string,
	ownerPart = HOST_PART,
): Promise<void> {
	const relsPath = relationshipsPartFor(ownerPart);
	const xml = await zip.file(relsPath)?.async('string');
	if (xml) {
		const document = parseXml(xml);
		const entry = elements(document.documentElement).find(
			(node) =>
				node.namespaceURI === NS.rels &&
				node.getAttribute('Id') === relId &&
				node.getAttribute('Type') === DIAGRAM_RELATIONSHIP_TYPES.drawing,
		);
		if (entry) {
			document.documentElement.removeChild(entry);
			zip.file(relsPath, buildXml(document));
		}
	}
	// Another relationship (from any part) may still use the drawing: then keep it.
	for (const file of zip.file(/\.rels$/)) {
		const sourcePart =
			file.name === '_rels/.rels' ? '' : file.name.replace(/(^|\/)_rels\/([^/]+)\.rels$/, '$1$2');
		try {
			for (const rel of parseRelationships(await file.async('string')).values())
				if (rel.mode !== 'External' && resolvePartPath(sourcePart, rel.target) === partName) return;
		} catch {
			return;
		}
	}
	zip.remove(partName);
	await removeContentTypeOverride(zip, partName);
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
		const reformatted =
			diagram.colorsId !== original.colorsId || diagram.quickStyleId !== original.quickStyleId;
		if (
			!relaid &&
			!reformatted &&
			JSON.stringify(nodesOf(diagram)) === JSON.stringify(nodesOf(original))
		)
			continue;
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
		for (const [key, xml] of [
			['colors', diagram.layoutSource?.colorsXml],
			['quickStyle', diagram.layoutSource?.quickStyleXml],
		] as const) {
			const name = diagram.parts[key]?.partName;
			if (reformatted && name && xml) zip.file(name, xml);
		}
		zip.file(
			partName,
			writeEditedDiagramData(xml, nodesOf(diagram), {
				dropDrawingLink: true,
				...(diagram.colorsId ? { colorsId: diagram.colorsId } : {}),
				...(diagram.quickStyleId ? { quickStyleId: diagram.quickStyleId } : {}),
				...(relaid && layoutId
					? { layout: { uniqueId: layoutId, category: diagram.layout?.family } }
					: {}),
			}),
		);
		if (diagram.layoutSource?.pictures?.length) {
			zip.file(
				partName,
				withDiagramPictureSources(await zip.file(partName)!.async('string'), diagram.layoutSource),
			);
			const relsName = relationshipsPartFor(partName);
			const rels = diagramPictureRelationships(
				partName,
				diagram.layoutSource,
				await zip.file(relsName)?.async('string'),
			);
			if (rels) zip.file(relsName, rels);
		}
		if (drawing) await dropDrawingPart(zip, drawing.relId, drawing.partName, drawing.ownerPart);
	}
}
