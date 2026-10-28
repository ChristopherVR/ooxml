import { NS, buildXml, first, type XmlElement } from '../xml/index';
import { parseDrawingColorIn } from './drawing-color';
import { readScene3DValues, readText3DValues, type ThreeDReader } from './read-3d-values';
import type { DrawingTextScene3D, DrawingTextShape3D, DrawingFlatText3D } from './text-3d-model';
export const DOM_3D_READER: ThreeDReader<XmlElement> = {
	attribute: (node, name) => node?.getAttribute(name),
	child: (node, name) => first(node, name, NS.a),
};
export function parseDrawingTextScene3D(node: XmlElement): DrawingTextScene3D {
	return { ...readScene3DValues(node, DOM_3D_READER), sourceXml: buildXml(node) };
}
export function parseDrawingTextShape3D(node: XmlElement): DrawingTextShape3D {
	const extrusionColor = parseDrawingColorIn(first(node, 'extrusionClr', NS.a));
	const contourColor = parseDrawingColorIn(first(node, 'contourClr', NS.a));
	return {
		...readText3DValues(node, DOM_3D_READER),
		sourceXml: buildXml(node),
		...(extrusionColor ? { extrusionColor } : {}),
		...(contourColor ? { contourColor } : {}),
	};
}
export function parseDrawingFlatText3D(node: XmlElement): DrawingFlatText3D {
	const raw = node.getAttribute('z');
	const z = raw === null || !raw.trim() ? NaN : Number(raw);
	return { sourceXml: buildXml(node), ...(Number.isSafeInteger(z) ? { z } : {}) };
}
