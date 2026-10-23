// Saving edited and inserted SmartArt in a worksheet drawing. An unedited diagram never reaches
// this module: its frame and parts are carried verbatim. An edited one keeps its frame XML (the
// anchor is rewritten) and its parts, except that its data part is rewritten by the shared
// writer (`writeEditedDiagramData`), a new layout replaces the layout part, and the cached drawing
// is dropped (relationship and part), because Excel ignores a stale drawing and lays the diagram
// out again (docs/agnostic-core-plan.md, "Excel and an edited data part"). An inserted one gets a
// `xdr:graphicFrame` with `dgm:relIds`, four new parts and their relationships, and no drawing.
import {
	DIAGRAM_CONTENT_TYPES,
	DIAGRAM_RELATIONSHIP_TYPES,
	withDiagramXmlProlog,
	writeEditedDiagramData,
	writeInsertedDiagramData,
} from '../../diagram/index';
import { NS } from '../../xml/index';
import type { DrawingObject, SmartArtObject } from '../model';
import type { PackageWriter, RelationshipSet } from './package-writer';
import { escapeAttr } from './xml-out';

type PartKey = 'data' | 'layout' | 'quickStyle' | 'colors';
const PART_KEYS: readonly PartKey[] = ['data', 'layout', 'quickStyle', 'colors'];
const FILE_NAMES: Record<PartKey, string> = {
	data: 'data',
	layout: 'layout',
	quickStyle: 'quickStyle',
	colors: 'colors',
};

const nodesOf = (object: SmartArtObject) =>
	object.nodes.map(({ id, text, parentId, nodeType }) => ({ id, text, parentId, nodeType }));

/** Whether the save writes this SmartArt's parts (edited or inserted) instead of carrying them. */
export const isRewrittenSmartArt = (object: DrawingObject): object is SmartArtObject =>
	object.kind === 'smartArt' && Boolean(object.edited || object.inserted);

/**
 * Package part names of the cached drawings edited diagrams drop. The drawing writer neither keeps
 * nor carries a relationship to one of them, unless an unedited diagram still uses it.
 */
export function droppedSmartArtDrawings(objects: readonly DrawingObject[]): Set<string> {
	const dropped = new Set<string>();
	const kept = new Set<string>();
	for (const object of objects) {
		if (object.kind !== 'smartArt' || !object.parts?.drawing) continue;
		(object.edited ? dropped : kept).add(object.parts.drawing);
	}
	for (const part of kept) dropped.delete(part);
	return dropped;
}

/** Rewrites the data part (and the layout part after Change Layout) of an edited diagram. */
export function writeEditedSmartArt(writer: PackageWriter, object: SmartArtObject): void {
	const dataPart = object.parts?.data;
	const xml = dataPart ? writer.source?.text(dataPart) : undefined;
	if (!dataPart || xml === undefined) return;
	const layoutPart = object.parts?.layout;
	const layoutXml = object.definitions?.layout;
	const relaid = Boolean(layoutXml && layoutPart && object.layoutId);
	if (relaid && layoutPart && layoutXml) {
		writer.carry(layoutPart);
		writer.add(layoutPart, withDiagramXmlProlog(layoutXml));
	}
	writer.carry(dataPart);
	writer.add(
		dataPart,
		writeEditedDiagramData(xml, nodesOf(object), {
			dropDrawingLink: true,
			...(relaid && object.layoutId
				? { layout: { uniqueId: object.layoutId, category: object.layoutFamily } }
				: {}),
		}),
	);
}

/** The first `n` for which none of the four part names is taken. */
function insertedPartNames(writer: PackageWriter): Record<PartKey, string> {
	const taken = (name: string) => writer.has(name) || Boolean(writer.source?.has(name));
	for (let n = 1; ; n++) {
		const names = Object.fromEntries(
			PART_KEYS.map((key) => [key, `xl/diagrams/${FILE_NAMES[key]}${n}.xml`]),
		) as Record<PartKey, string>;
		if (!PART_KEYS.some((key) => taken(names[key]))) return names;
	}
}

/**
 * Writes the four parts of an inserted diagram, adds their relationships to the drawing part's
 * set and returns the frame (`xdr:graphicFrame`) that names them.
 */
export function writeInsertedSmartArt(
	writer: PackageWriter,
	object: SmartArtObject,
	rels: RelationshipSet,
	relativeTarget: (toPart: string) => string,
	id: number,
): string {
	const names = insertedPartNames(writer);
	const definitions = object.definitions ?? {};
	const content: Record<PartKey, string | undefined> = {
		data: writeInsertedDiagramData({
			nodes: object.nodes,
			layoutId: object.layoutId ?? '',
			layoutCategory: object.layoutFamily,
			colorsId: object.colorsId ?? '',
			quickStyleId: object.quickStyleId ?? '',
		}),
		layout: definitions.layout,
		quickStyle: definitions.quickStyle,
		colors: definitions.colors,
	};
	const relIds = {} as Record<PartKey, string>;
	for (const key of PART_KEYS) {
		const xml = content[key];
		if (xml === undefined) throw new Error(`The inserted SmartArt has no ${key} definition.`);
		writer.add(names[key], withDiagramXmlProlog(xml), DIAGRAM_CONTENT_TYPES[key]);
		relIds[key] = rels.add(DIAGRAM_RELATIONSHIP_TYPES[key], relativeTarget(names[key]));
	}
	const name = escapeAttr(object.name ?? `Diagram ${id}`);
	return (
		`<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
		'<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
		`<a:graphic><a:graphicData uri="${NS.dgm}"><dgm:relIds xmlns:dgm="${NS.dgm}" r:dm="${relIds.data}" r:lo="${relIds.layout}" r:qs="${relIds.quickStyle}" r:cs="${relIds.colors}"/></a:graphicData></a:graphic></xdr:graphicFrame>`
	);
}
