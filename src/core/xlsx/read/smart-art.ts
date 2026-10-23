// SmartArt (DiagramML) in a worksheet drawing: the `xdr:graphicFrame` whose `a:graphicData` uses
// the diagram namespace. The frame is read synchronously with the drawing part; the five diagram
// parts are read afterwards through the format-neutral `diagram` area (`loadDiagram`), as
// `src/core/docx/diagram.ts` does for Word. Unedited parts are kept verbatim on save.
import {
	attributeReader,
	diagramNodesInTreeOrder,
	loadDiagram,
	parseDiagramDataModel,
	parseRelationshipIdAttributes,
	type DiagramRelationshipIds,
} from '../../diagram/index';
import { parseRelationships, resolvePartPath } from '../../opc/index';
import { NS, first } from '../../xml/index';
import type { DrawingAnchor, SmartArtObject, Worksheet } from '../model';
import type { SourceIndex } from './package';
import { SMART_ART_NOTICES } from '../smart-art-notices';

/** The `a:graphicData/@uri` of a SmartArt graphic. */
export const SMART_ART_GRAPHIC_URI = NS.dgm;

const NOTICES = SMART_ART_NOTICES;

interface Pending {
	hostPart: string;
	relationshipIds: DiagramRelationshipIds;
}

/** SmartArt objects read from a drawing whose diagram parts are not loaded yet. */
const pending = new WeakMap<SmartArtObject, Pending>();

/** The frame of a SmartArt graphic, before its diagram parts are read. */
export function smartArtFrame(
	graphicData: Element,
	anchor: DrawingAnchor,
	name: string | undefined,
	sourceXml: string,
	hostPart: string,
): SmartArtObject {
	const relIds = first(graphicData, 'relIds', NS.dgm);
	const relationshipIds = relIds ? parseRelationshipIdAttributes(attributeReader(relIds)) : {};
	const object: SmartArtObject = {
		kind: 'smartArt',
		anchor,
		...(name ? { name } : {}),
		nodes: [],
		notice: NOTICES.placeholder,
		issues: relIds
			? []
			: [
					{
						code: 'DIAGRAM_RELIDS_MISSING',
						message: 'The graphic has no dgm:relIds, so no diagram part can be located.',
					},
				],
		sourceXml,
	};
	pending.set(object, { hostPart, relationshipIds });
	return object;
}

/** Reads the diagram parts of one SmartArt object. Never throws: problems go to `issues`. */
async function resolveOne(object: SmartArtObject, info: Pending, source: SourceIndex) {
	const { hostPart, relationshipIds } = info;
	const rels = source.rels(hostPart);
	const own = new Set(Object.values(relationshipIds));
	// Only the cached drawing the data model names is offered: guessing among the sheet's drawing
	// relationships could attach another diagram's drawing.
	let declared: string | undefined;
	const dataRel = relationshipIds.dataRelId ? rels.get(relationshipIds.dataRelId) : undefined;
	if (dataRel) {
		try {
			const xml = source.text(resolvePartPath(hostPart, dataRel.target));
			declared = xml ? parseDiagramDataModel(xml).drawingRelId : undefined;
		} catch {
			// Reported by loadDiagram below.
		}
	}
	const offered = new Map(
		[...rels.values()]
			.filter((rel) => rel.mode !== 'External' && (own.has(rel.id) || rel.id === declared))
			.map((rel) => [rel.id, rel.target] as const),
	);
	const loaded = await loadDiagram(hostPart, relationshipIds, {
		hostRelationships: () => offered,
		resolvePath: resolvePartPath,
		readText: async (part) => source.text(part),
		readRelationshipEntries: async (relsPart) => {
			const xml = source.text(relsPart);
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
	object.issues.push(...loaded.issues);
	const paths = loaded.paths;
	object.parts = {
		...(paths.data ? { data: paths.data } : {}),
		...(paths.layout ? { layout: paths.layout } : {}),
		...(paths.quickStyle ? { quickStyle: paths.quickStyle } : {}),
		...(paths.colors ? { colors: paths.colors } : {}),
		...(paths.drawing ? { drawing: paths.drawing } : {}),
	};
	if (loaded.layout?.uniqueId) object.layoutId = loaded.layout.uniqueId;
	if (loaded.layout?.family) object.layoutFamily = loaded.layout.family;
	if (loaded.colors?.uniqueId) object.colorsId = loaded.colors.uniqueId;
	if (loaded.quickStyle?.uniqueId) object.quickStyleId = loaded.quickStyle.uniqueId;
	if (loaded.data) object.nodes = diagramNodesInTreeOrder(loaded.data);
	if (loaded.drawing && loaded.drawing.shapes.length > 0) {
		object.diagram = loaded.drawing;
		object.notice = NOTICES.cached;
	}
}

/**
 * Loads the diagram parts of every SmartArt object read from `sheets`' drawings: node text,
 * layout id and the cached drawing. Warns once per object that has no drawing to show.
 */
export async function resolveSmartArt(
	sheets: readonly Worksheet[],
	source: SourceIndex,
	warn: (message: string) => void,
): Promise<void> {
	for (const sheet of sheets)
		for (const drawing of sheet.drawings) {
			if (drawing.kind !== 'smartArt') continue;
			const info = pending.get(drawing);
			if (!info) continue;
			pending.delete(drawing);
			await resolveOne(drawing, info, source);
			if (!drawing.diagram)
				warn(
					`A SmartArt graphic on sheet "${sheet.name}" has no cached drawing; it is drawn with an approximate layout. It is kept on save.`,
				);
		}
}
