// Finds the cached `dsp:drawing` part of a diagram. Extracted from
// `pptx/core/core/runtime/smartart-drawing-part.ts` and made host-neutral: the host (a slide, a
// Word document, ...) supplies its relationship lookup instead of a presentation runtime.
import { RELATIONSHIP_TYPES } from '../opc/index.js';

/** One entry of a `.rels` part. */
export interface DiagramRelationshipEntry {
	id: string;
	type: string;
	target: string;
}

/** What the resolver needs from the package that hosts the diagram. */
export interface DiagramPartHost {
	/** Relationships (id -> target) of the part that holds the `dgm:relIds` (slide, `word/document.xml`). */
	hostRelationships(hostPart: string): ReadonlyMap<string, string> | undefined;
	/** Resolves a relationship target against the part that declared it. */
	resolvePath(basePart: string, target: string): string;
	/** Entries of a `.rels` part, or `undefined` when it is absent or unreadable; malformed XML is the host's to swallow. */
	readRelationshipEntries(relsPart: string): Promise<DiagramRelationshipEntry[] | undefined>;
}

/** `ppt/diagrams/data1.xml` -> `ppt/diagrams/_rels/data1.xml.rels`. */
export function diagramRelationshipsPart(partPath: string): string {
	const dir = partPath.replace(/\/[^/]+$/u, '');
	const file = partPath.split('/').pop() ?? '';
	return `${dir}/_rels/${file}.rels`;
}

/**
 * Resolves the cached drawing of a diagram, in the order producers use: the `dsp:dataModelExt`
 * relationship id on the host part, a drawing relationship the host carries without an extension,
 * then a `diagramDrawing` relationship on the data part itself (older files).
 */
export async function resolveDiagramDrawingPart(
	hostPart: string,
	dataRelId: string,
	drawingExtensionRelId: string,
	host: DiagramPartHost,
): Promise<{ relId: string; path: string } | undefined> {
	if (dataRelId.length === 0) return undefined;
	const hostRels = host.hostRelationships(hostPart);
	const declared = drawingExtensionRelId ? hostRels?.get(drawingExtensionRelId) : undefined;
	if (declared) return { relId: drawingExtensionRelId, path: host.resolvePath(hostPart, declared) };

	// Some producers omit dataModelExt but still leave a single drawing relationship on the host.
	const inferred = [...(hostRels?.entries() ?? [])].find(([, target]) =>
		/(?:^|\/)diagrams\/drawing\d+\.xml$/u.test(target.replaceAll('\\', '/')),
	);
	if (inferred) return { relId: inferred[0], path: host.resolvePath(hostPart, inferred[1]) };

	const dataTarget = hostRels?.get(dataRelId);
	if (!dataTarget) return undefined;
	const dataPath = host.resolvePath(hostPart, dataTarget);
	const entries = await host.readRelationshipEntries(diagramRelationshipsPart(dataPath));
	if (!entries) return undefined;
	const drawing = entries.find(
		(entry) =>
			(!drawingExtensionRelId || entry.id.trim() === drawingExtensionRelId) &&
			entry.type.endsWith('/diagramDrawing'),
	);
	const id = drawing?.id.trim() ?? '';
	const target = drawing?.target.trim() ?? '';
	if (id.length === 0 || target.length === 0) return undefined;
	return { relId: id, path: host.resolvePath(dataPath, target) };
}

/** Relationship types of the five diagram parts, in `dgm:relIds` order plus the cached drawing. */
export const DIAGRAM_RELATIONSHIP_TYPES = {
	data: RELATIONSHIP_TYPES.diagramData,
	layout: RELATIONSHIP_TYPES.diagramLayout,
	quickStyle: RELATIONSHIP_TYPES.diagramQuickStyle,
	colors: RELATIONSHIP_TYPES.diagramColors,
	drawing: RELATIONSHIP_TYPES.diagramDrawing,
} as const;
