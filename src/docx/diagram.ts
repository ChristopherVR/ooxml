// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// SmartArt (DiagramML) in a Word document: the `w:drawing` whose `a:graphicData` uses the diagram
// namespace. This module models the host side (extent, anchor/inline, relationships to the four
// diagram parts and the cached drawing) and reads the five parts through the format-neutral
// `diagram` area. The parts themselves are never rewritten: they stay in the package untouched.
import {
	DIAGRAM_RELATIONSHIP_TYPES,
	attributeReader,
	loadDiagram,
	parseDiagramDataModel,
	parseRelationshipIdAttributes,
	type DiagramDrawing,
	type DiagramIssue,
	type DiagramLayoutType,
	type DiagramRelationshipIds,
} from '../diagram/index.js';
import { parseRelationships, resolvePartPath } from '../opc/index.js';
import type { Relationship } from './package-parts.js';
import type { XmlElement } from './xml.js';

/** The `a:graphicData/@uri` of a SmartArt graphic. */
export const DIAGRAM_GRAPHIC_URI = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const DGM_NS = DIAGRAM_GRAPHIC_URI;

/** A package part a diagram refers to. */
export interface DocxDiagramPart {
	/** Relationship id in the part that holds the drawing (`word/document.xml`, a header...). */
	relId: string;
	/** Package part name, e.g. `word/diagrams/data1.xml`. */
	partName: string;
}

/** One content node of the diagram's data model. */
export interface DocxDiagramNode {
	id: string;
	text: string;
	/** Model id of the parent content node, when the data model has one. */
	parentId?: string;
}

/**
 * How the diagram can be shown.
 * `cached-drawing`: the shapes the producing application saved (`dsp:drawing`) are available as
 * `drawing`; they are a snapshot, not a layout computed here.
 * `placeholder`: no usable cached drawing; only a labelled placeholder (with the node text) can be
 * shown, because no SmartArt layout is computed.
 */
export type DocxDiagramRendering = 'cached-drawing' | 'placeholder';

/** A SmartArt graphic in a Word document. Part contents are read-only; the XML is preserved on save. */
export interface DocxDiagram {
	/** `inline` (`wp:inline`) or `anchor` (`wp:anchor`, floating). */
	placement: 'inline' | 'anchor';
	/** `wp:extent` in EMU: the frame the cached drawing's coordinates are relative to. */
	extentEmu: { width: number; height: number };
	/** `wp:docPr/@name` and `@id`. */
	name?: string;
	id?: string;
	/** `dgm:relIds` as written on the graphic. */
	relationshipIds: DiagramRelationshipIds;
	/** The diagram parts the relationships resolve to (package part names). */
	parts: {
		data?: DocxDiagramPart;
		layout?: DocxDiagramPart;
		quickStyle?: DocxDiagramPart;
		colors?: DocxDiagramPart;
		/** Cached layout (`dsp:drawing`), once located. */
		drawing?: DocxDiagramPart;
	};
	/**
	 * Relationship id -> package part name for the relationships this diagram needs, in the part
	 * that holds it. Lets the parts be resolved later without the host part's `.rels`.
	 */
	hostRelationships: Record<string, string>;
	/** Layout identification, from `dgm:layoutDef`. */
	layout?: { uniqueId: string; family?: DiagramLayoutType; rootAlgorithm?: string };
	/** Colour and quick-style definition ids (`dgm:colorsDef/@uniqueId`, `dgm:styleDef/@uniqueId`). */
	colorsId?: string;
	quickStyleId?: string;
	/** Content nodes of the data model, in document order. */
	nodes: DocxDiagramNode[];
	/** The cached shape tree, EMU relative to `extentEmu`. Absent for `placeholder` rendering. */
	drawing?: DiagramDrawing;
	/** Whether a renderer can draw shapes (`cached-drawing`) or must show a placeholder. */
	rendering: DocxDiagramRendering;
	/** What this model does and does not provide, to show alongside the diagram. */
	notice: string;
	/** Problems met reading the parts; empty when everything resolved. */
	issues: DiagramIssue[];
}

const emu = (element: XmlElement | undefined, name: string): number => {
	const value = Number(element?.getAttribute(name));
	return Number.isFinite(value) ? value : 0;
};

const NOTICES = {
	cached:
		'Shown from the drawing cached in the file (the layout last saved by the producing application). Word layout is not recomputed, so the diagram can differ from Word after edits elsewhere, and its content is not editable here; the diagram parts are preserved unchanged on save.',
	placeholder:
		'This file has no usable cached drawing for the diagram, and SmartArt layout is not computed here, so only a placeholder with the node text is shown. The diagram parts are preserved unchanged on save.',
} as const;

const isDrawingRelationship = (diagram: DocxDiagram, relId: string): boolean =>
	!Object.values(diagram.relationshipIds).includes(relId);

/** The notice that goes with a rendering. */
export const diagramNotice = (rendering: DocxDiagramRendering): string =>
	rendering === 'cached-drawing' ? NOTICES.cached : NOTICES.placeholder;

/**
 * Reads the host side of a SmartArt graphic. Synchronous: resolves the relationships through the
 * host's `.rels` but reads no part; {@link resolveDiagramParts} fills in the part contents.
 */
export function parseDiagramGraphic(
	graphicData: XmlElement,
	extent: XmlElement | undefined,
	docPr: XmlElement | undefined,
	rels: ReadonlyMap<string, Relationship>,
	isAnchor: boolean,
): DocxDiagram {
	const relIds = Array.from(graphicData.getElementsByTagNameNS(DGM_NS, 'relIds'))[0];
	const relationshipIds = relIds ? parseRelationshipIdAttributes(attributeReader(relIds)) : {};
	const hostRelationships: Record<string, string> = {};
	const issues: DiagramIssue[] = [];
	const part = (relId: string | undefined, type: string): DocxDiagramPart | undefined => {
		if (!relId) return undefined;
		const rel = rels.get(relId);
		if (!rel || rel.mode === 'External') {
			issues.push({
				code: 'DIAGRAM_RELATIONSHIP_UNRESOLVED',
				message: `Diagram relationship ${relId} does not point at a part in this package.`,
			});
			return undefined;
		}
		if (rel.type !== type)
			issues.push({
				code: 'DIAGRAM_RELATIONSHIP_TYPE',
				message: `Diagram relationship ${relId} has type ${rel.type}, expected ${type}.`,
			});
		const partName = resolvePartPath('word/document.xml', rel.target);
		hostRelationships[relId] = partName;
		return { relId, partName };
	};
	const parts: DocxDiagram['parts'] = {};
	const slots = [
		['data', relationshipIds.dataRelId, DIAGRAM_RELATIONSHIP_TYPES.data],
		['layout', relationshipIds.layoutRelId, DIAGRAM_RELATIONSHIP_TYPES.layout],
		['quickStyle', relationshipIds.styleRelId, DIAGRAM_RELATIONSHIP_TYPES.quickStyle],
		['colors', relationshipIds.colorsRelId, DIAGRAM_RELATIONSHIP_TYPES.colors],
	] as const;
	for (const [name, relId, type] of slots) {
		const resolved = part(relId, type);
		if (resolved) parts[name] = resolved;
	}
	if (!relIds)
		issues.push({
			code: 'DIAGRAM_RELIDS_MISSING',
			message: 'The graphic has no dgm:relIds, so no diagram part can be located.',
		});
	// Every diagramDrawing relationship of the host: the cached drawing is found among them.
	for (const rel of rels.values())
		if (rel.type === DIAGRAM_RELATIONSHIP_TYPES.drawing && rel.mode !== 'External')
			hostRelationships[rel.id] = resolvePartPath('word/document.xml', rel.target);
	const name = docPr?.getAttribute('name') || undefined;
	const id = docPr?.getAttribute('id') || undefined;
	return {
		placement: isAnchor ? 'anchor' : 'inline',
		extentEmu: { width: emu(extent, 'cx'), height: emu(extent, 'cy') },
		...(name ? { name } : {}),
		...(id ? { id } : {}),
		relationshipIds,
		parts,
		hostRelationships,
		nodes: [],
		rendering: 'placeholder',
		notice: diagramNotice('placeholder'),
		issues,
	};
}

/**
 * Reads the diagram's parts from the package and completes the model: node text, layout and style
 * identification, the cached drawing and the honest rendering status. Never throws; problems go to
 * `issues`.
 */
export async function resolveDiagramParts(
	diagram: DocxDiagram,
	readText: (partName: string) => Promise<string | undefined>,
): Promise<void> {
	const hostPart = 'word/document.xml';
	// Only the drawing the data model names is offered to the loader: guessing among a document's
	// drawing relationships could attach another diagram's drawing.
	const dataPart = diagram.parts.data?.partName;
	let declaredDrawing: string | undefined;
	try {
		const dataXml = dataPart ? await readText(dataPart) : undefined;
		declaredDrawing = dataXml ? parseDiagramDataModel(dataXml).drawingRelId : undefined;
	} catch {
		// Reported by loadDiagram below.
	}
	const offered = Object.entries(diagram.hostRelationships).filter(
		([id]) => id === declaredDrawing || !isDrawingRelationship(diagram, id),
	);
	const loaded = await loadDiagram(hostPart, diagram.relationshipIds, {
		hostRelationships: () => new Map(offered.map(([id, name]) => [id, `/${name}`])),
		resolvePath: resolvePartPath,
		readText,
		readRelationshipEntries: async (relsPart) => {
			const xml = await readText(relsPart);
			if (!xml) return undefined;
			try {
				return [...parseRelationships(xml).values()].map(({ id, type, target }) => ({
					id,
					type,
					target,
				}));
			} catch {
				return undefined;
			}
		},
	});
	// Relationship problems were already recorded while parsing the host markup.
	const known = new Set(diagram.issues.map((issue) => `${issue.code}|${issue.message}`));
	const fresh = loaded.issues.filter(
		(issue) =>
			!issue.code.startsWith('DIAGRAM_RELATIONSHIP_') &&
			!known.has(`${issue.code}|${issue.message}`),
	);
	diagram.issues.push(...fresh);
	if (loaded.paths.drawing && loaded.drawingRelId)
		diagram.parts.drawing = { relId: loaded.drawingRelId, partName: loaded.paths.drawing };
	if (loaded.layout) {
		diagram.layout = {
			uniqueId: loaded.layout.uniqueId,
			...(loaded.layout.family ? { family: loaded.layout.family } : {}),
			...(loaded.layout.rootAlgorithm ? { rootAlgorithm: loaded.layout.rootAlgorithm } : {}),
		};
	}
	if (loaded.colors?.uniqueId) diagram.colorsId = loaded.colors.uniqueId;
	if (loaded.quickStyle?.uniqueId) diagram.quickStyleId = loaded.quickStyle.uniqueId;
	if (loaded.data)
		diagram.nodes = loaded.data.nodes.map((node) => ({
			id: node.modelId,
			text: node.text,
			...(loaded.data?.parentById.get(node.modelId)
				? { parentId: loaded.data.parentById.get(node.modelId) as string }
				: {}),
		}));
	if (loaded.drawing && loaded.drawing.shapes.length > 0) {
		diagram.drawing = loaded.drawing;
		diagram.rendering = 'cached-drawing';
	} else {
		delete diagram.drawing;
		diagram.rendering = 'placeholder';
	}
	diagram.notice = diagramNotice(diagram.rendering);
}
