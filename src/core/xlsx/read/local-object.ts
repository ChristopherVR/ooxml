import { MODEL3D_REL, MODEL3D_URI } from '../../drawingml/index';
import { embeddableOfficePackageFormat, RELATIONSHIP_TYPES } from '../../opc/index';
import { NS, elements, first, relAttr, type XmlElement } from '../../xml/index';
import type { DrawingAnchor, EmbeddedDocumentObject, Model3DObject } from '../model';
import type { SourceIndex } from './package';
import { att, selfContainedXml } from './xml-util';

function relatedPart(source: SourceIndex, part: string, id: string | undefined, type: string) {
	const rel = id ? source.rels(part).get(id) : undefined;
	const target = rel?.type === type ? source.target(part, rel) : undefined;
	return target && source.has(target) ? target : undefined;
}

/** Recognizes the native choice and its raster fallback while retaining the whole anchor. */
export function readModel3D(
	source: SourceIndex,
	part: string,
	content: XmlElement,
	anchorNode: XmlElement,
	anchor: DrawingAnchor,
): Model3DObject | undefined {
	const frame = first(first(content, 'Choice', NS.mc), 'graphicFrame', NS.xdr);
	const data = first(first(frame, 'graphic', NS.a), 'graphicData', NS.a);
	if (att(data, 'uri') !== MODEL3D_URI) return undefined;
	const model = first(data, 'model3d', MODEL3D_URI);
	const partName = relatedPart(source, part, relAttr(model, 'embed'), MODEL3D_REL);
	const fallback = first(first(content, 'Fallback', NS.mc), 'pic', NS.xdr);
	const picture = first(first(fallback, 'blipFill', NS.xdr), 'blip', NS.a);
	const raster = first(first(model, 'raster', MODEL3D_URI), 'blip', MODEL3D_URI);
	const previewPartName =
		relatedPart(source, part, relAttr(picture, 'embed'), RELATIONSHIP_TYPES.image) ??
		relatedPart(source, part, relAttr(raster, 'embed'), RELATIONSHIP_TYPES.image);
	const previewContentType = previewPartName ? source.contentType(previewPartName) : undefined;
	if (!partName || !previewPartName || !previewContentType?.startsWith('image/')) return undefined;
	const name = att(first(first(frame, 'nvGraphicFramePr', NS.xdr), 'cNvPr', NS.xdr), 'name');
	return {
		kind: 'model3d',
		name: name ?? '',
		anchor,
		partName,
		previewPartName,
		previewContentType,
		sourceXml: selfContainedXml(anchorNode),
	};
}

interface EmbeddedDeclaration {
	node: XmlElement;
	partName: string;
	contentType: string;
	progId: string;
	previewPartName: string;
	relationIds: string[];
}

export interface EmbeddedObjectDeclarations {
	found: ReadonlyMap<number, EmbeddedDeclaration>;
	read: Set<EmbeddedDeclaration>;
}

/** Finds editable, picture-backed Office packages. Other worksheet OLE entries remain intact. */
export function readEmbeddedDeclarations(
	source: SourceIndex,
	sheetPart: string,
	node: XmlElement,
): EmbeddedObjectDeclarations {
	const found = new Map<number, EmbeddedDeclaration>();
	for (const child of elements(node)) {
		if (child.localName !== 'oleObject' || child.namespaceURI !== NS.x) continue;
		const id = relAttr(child, 'id');
		const previewId = relAttr(first(child, 'objectPr', NS.x), 'id');
		const shapeId = Number(att(child, 'shapeId'));
		const progId = att(child, 'progId');
		const partName = relatedPart(source, sheetPart, id, RELATIONSHIP_TYPES.package);
		const previewPartName = relatedPart(source, sheetPart, previewId, RELATIONSHIP_TYPES.image);
		const contentType = partName ? source.contentType(partName) : undefined;
		if (
			att(child, 'dvAspect') !== 'DVASPECT_ICON' ||
			att(child, 'link') !== undefined ||
			!Number.isSafeInteger(shapeId) ||
			shapeId < 1 ||
			!id ||
			!previewId ||
			!progId ||
			!contentType ||
			!partName ||
			!previewPartName
		)
			continue;
		const format = embeddableOfficePackageFormat(partName.split('.').pop() ?? '');
		const progMatches =
			format?.progId === progId || (format?.type === 'excel' && progId === 'Worksheet');
		if (!format || format.mime !== contentType || !progMatches) continue;
		found.set(shapeId, {
			node: child,
			partName,
			contentType,
			progId,
			previewPartName,
			relationIds: [id, previewId],
		});
	}
	return { found, read: new Set() };
}

export function embeddedPicture(
	picture: import('../model').ImageObject,
	id: string | undefined,
	objects: EmbeddedObjectDeclarations | undefined,
): EmbeddedDocumentObject | undefined {
	const found = objects?.found.get(Number(id));
	if (!found || found.previewPartName !== picture.partName) return undefined;
	objects?.read.add(found);
	return {
		kind: 'embeddedObject',
		name: picture.name ?? found.partName.split('/').pop() ?? '',
		anchor: picture.anchor,
		partName: found.partName,
		progId: found.progId,
		contentType: found.contentType,
		previewPartName: found.previewPartName,
		previewContentType: picture.contentType,
		...(picture.sourceXml ? { sourceXml: picture.sourceXml } : {}),
	};
}
