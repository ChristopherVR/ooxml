import {
	GLB_PART_CONTENT_TYPE,
	MODEL3D_REL,
	MODEL3D_URI,
	defaultModel3DSceneXml,
} from '../../drawingml/index';
import { RELATIONSHIP_TYPES } from '../../opc/index';
import { NS, elements, parseXml } from '../../xml/index';
import { anchorToPixelBox, pixelBoxToAnchor, pixelSizeToExtent } from '../layout/anchors';
import { createGridMetrics } from '../layout/metrics';
import type { DrawingAnchor, EmbeddedDocumentObject, Model3DObject, Worksheet } from '../model';
import type { PackageWriter, RelationshipSet } from './package-writer';
import { escapeAttr, inlineFragment } from './xml-out';
import { selfContainedXml } from '../read/xml-util';

export function localPartRelationship(
	writer: PackageWriter,
	rels: RelationshipSet,
	relative: (target: string) => string,
	asset: { name: string; mime: string; relation: string },
): string | undefined {
	const bytes = writer.source?.bytes(asset.name);
	if (!bytes) return undefined;
	if (!writer.has(asset.name)) writer.add(asset.name, bytes, asset.mime);
	return rels.ensure(asset.relation, relative(asset.name));
}

type Marker = DrawingAnchor['from'];

function objectMarkerXml(marker: Marker): string {
	return `<xdr:col>${marker.col}</xdr:col><xdr:colOff>${Math.round(marker.colOffset)}</xdr:colOff><xdr:row>${marker.row}</xdr:row><xdr:rowOff>${Math.round(marker.rowOffset)}</xdr:rowOff>`;
}

/** The modern worksheet OLE record references both the embedded package and the preview. */
export function embeddedObjectXml(
	writer: PackageWriter,
	rels: RelationshipSet,
	relative: (target: string) => string,
	sheet: Worksheet,
	object: EmbeddedDocumentObject,
	id: number,
): string | undefined {
	const packageId = localPartRelationship(writer, rels, relative, {
		name: object.partName,
		mime: object.contentType,
		relation: RELATIONSHIP_TYPES.package,
	});
	const previewId = localPartRelationship(writer, rels, relative, {
		name: object.previewPartName,
		mime: object.previewContentType,
		relation: RELATIONSHIP_TYPES.image,
	});
	if (!packageId || !previewId) return undefined;
	const metrics = createGridMetrics(sheet, { zoom: 100 });
	const { from, to } = object.anchor.to
		? object.anchor
		: pixelBoxToAnchor(sheet, anchorToPixelBox(sheet, object.anchor, metrics), metrics, 'twoCell');
	if (!to) return undefined;
	return (
		`<oleObject progId="${escapeAttr(object.progId)}" dvAspect="DVASPECT_ICON" shapeId="${id}" r:id="${packageId}">` +
		`<objectPr defaultSize="0" r:id="${previewId}"><anchor moveWithCells="1">` +
		`<from>${objectMarkerXml(from)}</from><to>${objectMarkerXml(to)}</to>` +
		'</anchor></objectPr></oleObject>'
	);
}

/** Combines newly written entries and preserved linked, legacy or unmodeled objects. */
export function mergeOleObjects(fresh: readonly string[], kept: readonly string[]): string {
	if (!fresh.length) return '';
	const entries = kept.flatMap((xml) => {
		const root = parseXml(xml, { label: 'XLSX OLE objects' }).documentElement;
		return elements(root).map((entry) => inlineFragment(selfContainedXml(entry)));
	});
	return `<oleObjects xmlns:xdr="${NS.xdr}">${entries.join('')}${fresh.join('')}</oleObjects>`;
}

/** The Office model choice uses the same raster part as its DrawingML picture fallback. */
export function model3DFrame(
	model: Model3DObject,
	id: number,
	modelId: string,
	previewId: string,
	picture: string,
	sheet: Worksheet | undefined,
): string {
	const box = sheet
		? anchorToPixelBox(sheet, model.anchor, createGridMetrics(sheet, { zoom: 100 }))
		: undefined;
	const extent = box
		? pixelSizeToExtent(box.w, box.h)
		: (model.anchor.ext ?? pixelSizeToExtent(180, 180));
	return (
		`<mc:AlternateContent xmlns:mc="${NS.mc}"><mc:Choice xmlns:am3d="${MODEL3D_URI}" Requires="am3d">` +
		`<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${escapeAttr(model.name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
		'<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
		`<a:graphic><a:graphicData uri="${MODEL3D_URI}"><am3d:model3d r:embed="${modelId}">` +
		`<am3d:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${Math.round(extent.cx)}" cy="${Math.round(extent.cy)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></am3d:spPr>` +
		defaultModel3DSceneXml(previewId) +
		`</am3d:model3d></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback>${picture}</mc:Fallback></mc:AlternateContent>`
	);
}

export const modelPartRelationship = (object: Model3DObject) => ({
	name: object.partName,
	mime: GLB_PART_CONTENT_TYPE,
	relation: MODEL3D_REL,
});

export const previewPartRelationship = (object: Model3DObject | EmbeddedDocumentObject) => ({
	name: object.previewPartName,
	mime: object.previewContentType,
	relation: RELATIONSHIP_TYPES.image,
});
