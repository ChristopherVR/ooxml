import { MODEL3D_REL, MODEL3D_URI } from '../drawingml/model3d';
import { RELATIONSHIP_TYPES } from '../opc/index';
import { NS, first as firstChild } from '../xml/index';
import type { DrawingContext } from './drawing';
import type { InlineDrawingPayload, InlineImage } from './inline-model';
import { contentTypeForPart, resolveInternalTarget } from './package-parts';
import { getR, getW, type XmlElement } from './xml';

export const VML_NS = 'urn:schemas-microsoft-com:vml';
export const OFFICE_VML_NS = 'urn:schemas-microsoft-com:office:office';

function partFor(
	context: DrawingContext,
	relId: string | undefined,
	relationshipType: string,
): string | undefined {
	const rel = relId && context.rels.get(relId);
	if (!rel || rel.mode === 'External' || rel.type !== relationshipType) return undefined;
	const part = resolveInternalTarget(rel.target);
	return context.mediaParts.has(part) ? part : undefined;
}

function previewFor(
	context: DrawingContext,
	relId?: string,
): Pick<InlineImage, 'relId' | 'partName' | 'contentType'> | undefined {
	const partName = partFor(context, relId, RELATIONSHIP_TYPES.image);
	if (!partName || !relId) return undefined;
	const contentType = contentTypeForPart(context.contentTypes, partName)?.toLowerCase();
	if (!contentType || !/^image\/(png|jpeg|gif|bmp|webp)$/.test(contentType)) return undefined;
	return { relId, partName, contentType };
}

/** A model's renderer-supplied raster, which is also used by Word's compatibility picture. */
export function model3dDrawing(
	graphicData: XmlElement,
	context: DrawingContext,
): Pick<InlineImage, 'relId' | 'partName' | 'contentType' | 'payload' | 'unsupported'> {
	const model = firstChild(graphicData, 'model3d', MODEL3D_URI);
	const relId = getR(model, 'embed');
	const partName = partFor(context, relId, MODEL3D_REL);
	const raster = firstChild(model, 'raster', MODEL3D_URI);
	const blip = firstChild(raster, 'blip', MODEL3D_URI);
	const preview = previewFor(context, getR(blip, 'embed'));
	return {
		...(preview ?? {
			relId: relId ?? '',
			partName: '',
			contentType: 'application/octet-stream',
			unsupported: '3D Model',
		}),
		...(partName ? { payload: { kind: 'model3d', partName } } : {}),
	};
}

function sizePx(style: string, axis: 'width' | 'height', twips: string | undefined): number {
	const match = new RegExp(`(?:^|;)\\s*${axis}:\\s*([\\d.]+)pt`, 'i').exec(style);
	const points = match ? Number(match[1]) : Number(twips) / 20;
	return Number.isFinite(points) && points > 0 ? Math.round((points * 96) / 72) : 96;
}

/** Read an OLE icon for display and preserve its source markup, without activating any object. */
export function parseOleDrawing(node: XmlElement, context: DrawingContext): InlineImage {
	const shape = firstChild(node, 'shape', VML_NS);
	const object = firstChild(node, 'OLEObject', OFFICE_VML_NS);
	const localRelId = getR(object, 'id');
	const partName = partFor(context, localRelId, RELATIONSHIP_TYPES.package);
	const progId = object?.getAttribute('ProgID');
	const payload: InlineDrawingPayload | undefined =
		partName && progId && object?.getAttribute('Type') === 'Embed'
			? { kind: 'officePackage', partName, progId }
			: undefined;
	const imagedata = firstChild(shape, 'imagedata', VML_NS);
	const preview = previewFor(context, getR(imagedata, 'id'));
	const style = shape?.getAttribute('style') ?? '';
	const title = shape?.getAttribute('alt') || imagedata?.getAttributeNS(OFFICE_VML_NS, 'title');
	return {
		...(preview ?? {
			relId: localRelId ?? '',
			partName: '',
			contentType: 'application/octet-stream',
		}),
		widthPx: sizePx(style, 'width', getW(node, 'dxaOrig')),
		heightPx: sizePx(style, 'height', getW(node, 'dyaOrig')),
		...(title ? { altText: title, title } : {}),
		...(preview ? {} : { unsupported: 'Embedded object' }),
		...(payload ? { payload } : {}),
	};
}

export function isOleDrawing(node: XmlElement): boolean {
	return node.namespaceURI === NS.w && node.localName === 'object';
}
