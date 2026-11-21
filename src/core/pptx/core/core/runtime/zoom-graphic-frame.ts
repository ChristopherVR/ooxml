import type { XmlObject, ZoomPptxElement } from '../../types';
import { ensureXmlChild, ensureXmlChildOrCreate } from '../../utils/xml-access';

export const ZOOM_FRAME_TAGS = {
	'pslz:sldZm': 'http://schemas.microsoft.com/office/powerpoint/2016/slidezoom',
	'psezm:sectionZm': 'http://schemas.microsoft.com/office/powerpoint/2016/sectionzoom',
	'psuz:summaryZm': 'http://schemas.microsoft.com/office/powerpoint/2016/summaryzoom',
} as const;

type ZoomTag = keyof typeof ZOOM_FRAME_TAGS;

export function readZoomFrame(frame: XmlObject | undefined) {
	const graphic = ensureXmlChild(frame, 'a:graphic');
	const data = ensureXmlChild(graphic, 'a:graphicData');
	if (!data) return undefined;
	for (const tag of Object.keys(ZOOM_FRAME_TAGS) as ZoomTag[]) {
		if (data['@_uri'] !== ZOOM_FRAME_TAGS[tag]) continue;
		const node = ensureXmlChild(data, tag);
		if (node) return { tag, node };
	}
	return undefined;
}

/** Native Zoom content uses frame-relative geometry, like other graphic data. */
export function readZoomFrameGeometry(frame: XmlObject, zoom: ZoomPptxElement, emu: number) {
	const transform = ensureXmlChild(frame, 'p:xfrm');
	const off = ensureXmlChild(transform, 'a:off');
	const ext = ensureXmlChild(transform, 'a:ext');
	const x = Number(off?.['@_x'] ?? 0) / emu;
	const y = Number(off?.['@_y'] ?? 0) / emu;
	const props = ensureXmlChild(ensureXmlChild(frame, 'p:nvGraphicFramePr'), 'p:cNvPr');
	return {
		...zoom,
		x,
		y,
		width: ext ? Number(ext['@_cx']) / emu : zoom.width,
		height: ext ? Number(ext['@_cy']) / emu : zoom.height,
		shapeId: props?.['@_id'] === undefined ? zoom.shapeId : String(props['@_id']),
		name: typeof props?.['@_name'] === 'string' ? props['@_name'] : zoom.name,
		altText: typeof props?.['@_descr'] === 'string' ? props['@_descr'] : zoom.altText,
		summaryTargets: zoom.summaryTargets?.map((target) => ({
			...target,
			x: target.x + x,
			y: target.y + y,
		})),
		rawXml: frame,
	};
}

export function writeZoomFrameGeometry(
	frame: XmlObject,
	zoom: XmlObject,
	tag: ZoomTag,
	transform: XmlObject,
): void {
	const off = ensureXmlChildOrCreate(transform, 'a:off');
	frame['p:xfrm'] = { 'a:off': off, 'a:ext': ensureXmlChildOrCreate(transform, 'a:ext') };
	const [prefix, local] = tag.split(':');
	const objectTag = `${prefix}:${local}Obj`;
	for (const object of Array.isArray(zoom[objectTag]) ? zoom[objectTag] : [zoom[objectTag]]) {
		const properties = ensureXmlChild(object, `${prefix}:zmPr`);
		const shape = ensureXmlChild(properties, 'p166:spPr');
		const child = ensureXmlChild(ensureXmlChild(shape, 'a:xfrm'), 'a:off');
		if (!child) continue;
		child['@_x'] = String(Number(child['@_x']) - Number(off['@_x']));
		child['@_y'] = String(Number(child['@_y']) - Number(off['@_y']));
	}
}

export function buildZoomFrame(tag: ZoomTag, zoom: XmlObject, fallback: XmlObject): XmlObject {
	const summary = tag === 'psuz:summaryZm';
	const nv = ensureXmlChild(fallback, summary ? 'p:nvGrpSpPr' : 'p:nvPicPr');
	const shape = ensureXmlChild(fallback, summary ? 'p:grpSpPr' : 'p:spPr');
	const transform = ensureXmlChild(shape, 'a:xfrm');
	const data: XmlObject = { '@_uri': ZOOM_FRAME_TAGS[tag] };
	data[tag] = zoom;
	const frame: XmlObject = {
		'p:nvGraphicFramePr': {
			'p:cNvPr': { ...ensureXmlChild(nv, 'p:cNvPr') },
			'p:cNvGraphicFramePr': { 'a:graphicFrameLocks': { '@_noChangeAspect': '1' } },
			'p:nvPr': {},
		},
		'p:xfrm': {},
		'a:graphic': { 'a:graphicData': data },
	};
	if (transform) writeZoomFrameGeometry(frame, zoom, tag, transform);
	return frame;
}
