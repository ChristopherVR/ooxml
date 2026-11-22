import { defaultModel3DSceneXml, MODEL3D_URI } from '../drawingml/model3d';
import { NS } from '../xml/index';
import { OFFICE_VML_NS, VML_NS } from './drawing-payload';
import type { InlineImage } from './inline-model';
import { EMU_PER_PIXEL } from './units';
import {
	importElement,
	makeNS,
	makeW,
	parseXml,
	REL_NS,
	WORD_NS,
	type XmlDocument,
	type XmlElement,
} from './xml';

const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const XMLNS = 'http://www.w3.org/2000/xmlns/';

/** Word saves the native 3D model and the raster as alternatives in the same run. */
export function model3dAlternateContent(
	doc: XmlDocument,
	fallbackDrawing: XmlElement,
	modelId: string,
	previewId: string,
	image: InlineImage,
): XmlElement {
	const nativeDrawing = importElement(doc, fallbackDrawing);
	const graphicData = nativeDrawing.getElementsByTagNameNS(NS.a, 'graphicData')[0];
	if (!graphicData) throw new Error('Expected DrawingML picture for the 3D model fallback');
	const cx = Math.max(1, Math.round(image.widthPx * EMU_PER_PIXEL));
	const cy = Math.max(1, Math.round(image.heightPx * EMU_PER_PIXEL));
	const model = parseXml(
		`<am3d:model3d xmlns:am3d="${MODEL3D_URI}" xmlns:a="${NS.a}" xmlns:r="${REL_NS}" r:embed="${modelId}">` +
			`<am3d:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
			'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></am3d:spPr>' +
			defaultModel3DSceneXml(previewId) +
			'</am3d:model3d>',
	).documentElement;
	if (!model) throw new Error('Expected a 3D model element');
	while (graphicData.firstChild) graphicData.removeChild(graphicData.firstChild);
	graphicData.setAttribute('uri', MODEL3D_URI);
	graphicData.appendChild(importElement(doc, model));
	const alternate = makeNS(doc, MC_NS, 'mc:AlternateContent');
	const choice = makeNS(doc, MC_NS, 'mc:Choice');
	choice.setAttributeNS(XMLNS, 'xmlns:am3d', MODEL3D_URI);
	choice.setAttribute('Requires', 'am3d');
	choice.appendChild(nativeDrawing);
	alternate.appendChild(choice);
	const fallback = makeNS(doc, MC_NS, 'mc:Fallback');
	fallback.appendChild(fallbackDrawing);
	alternate.appendChild(fallback);
	return alternate;
}

/** A native embedded Office OPC package, displayed as a VML picture icon. */
export function officePackageDrawing(
	doc: XmlDocument,
	image: InlineImage,
	progId: string,
	packageId: string,
	previewId: string,
	id: string,
): XmlElement {
	const object = makeW(doc, 'object');
	const widthPt = image.widthPx * 0.75;
	const heightPt = image.heightPx * 0.75;
	object.setAttributeNS(WORD_NS, 'w:dxaOrig', String(Math.round(widthPt * 20)));
	object.setAttributeNS(WORD_NS, 'w:dyaOrig', String(Math.round(heightPt * 20)));
	const shapeId = `_x0000_i${1024 + Number(id)}`;
	const template = makeNS(doc, VML_NS, 'v:shapetype');
	const templateId = `_x0000_t75_${id}`;
	template.setAttribute('id', templateId);
	template.setAttribute('coordsize', '21600,21600');
	template.setAttributeNS(OFFICE_VML_NS, 'o:spt', '75');
	template.setAttribute('filled', 'f');
	template.setAttribute('stroked', 'f');
	template.setAttribute('path', 'm0,0l0,21600,21600,21600,21600,0xe');
	object.appendChild(template);
	const shape = makeNS(doc, VML_NS, 'v:shape');
	shape.setAttribute('id', shapeId);
	shape.setAttribute('type', `#${templateId}`);
	shape.setAttribute('style', `width:${widthPt}pt;height:${heightPt}pt`);
	if (image.title) shape.setAttribute('alt', image.title);
	shape.setAttributeNS(OFFICE_VML_NS, 'o:ole', '');
	const imagedata = makeNS(doc, VML_NS, 'v:imagedata');
	imagedata.setAttributeNS(REL_NS, 'r:id', previewId);
	shape.appendChild(imagedata);
	object.appendChild(shape);
	const ole = makeNS(doc, OFFICE_VML_NS, 'o:OLEObject');
	ole.setAttribute('Type', 'Embed');
	ole.setAttribute('ProgID', progId);
	ole.setAttribute('ShapeID', shapeId);
	ole.setAttribute('DrawAspect', 'Icon');
	ole.setAttribute('ObjectID', `_${1024 + Number(id)}`);
	ole.setAttributeNS(REL_NS, 'r:id', packageId);
	object.appendChild(ole);
	return object;
}

/** Patch an existing OLE icon's frame without rewriting its VML source or activating its object. */
export function updateOleIconSize(object: XmlElement, widthPx: number, heightPx: number): void {
	const shape = object.getElementsByTagNameNS(VML_NS, 'shape')[0];
	if (!shape) return;
	const widthPt = widthPx * 0.75;
	const heightPt = heightPx * 0.75;
	const style = (shape.getAttribute('style') ?? '')
		.split(';')
		.filter((part) => part && !/^\s*(width|height)\s*:/i.test(part));
	shape.setAttribute('style', [...style, `width:${widthPt}pt`, `height:${heightPt}pt`].join(';'));
	object.setAttributeNS(WORD_NS, 'w:dxaOrig', String(Math.round(widthPt * 20)));
	object.setAttributeNS(WORD_NS, 'w:dyaOrig', String(Math.round(heightPt * 20)));
}
