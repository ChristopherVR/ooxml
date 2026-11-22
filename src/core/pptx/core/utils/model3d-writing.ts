import type { XmlObject } from '../types';

export const MODEL3D_URI = 'http://schemas.microsoft.com/office/drawing/2017/model3d';
export const MODEL3D_REL = 'http://schemas.microsoft.com/office/2017/06/relationships/model3d';

export function readModel3DGraphic(frame: XmlObject | undefined): XmlObject | undefined {
	const graphic = frame?.['a:graphic'] as XmlObject | undefined;
	const data = graphic?.['a:graphicData'] as XmlObject | undefined;
	return data?.['@_uri'] === MODEL3D_URI
		? (data['am3d:model3d'] as XmlObject | undefined)
		: undefined;
}

/** Default camera and scene, using the public MS-ODRAWXML CT_Model3D child order. */
export function defaultModel3DSceneXml(posterRid?: string): string {
	const raster = posterRid
		? `<am3d:raster rName="ooxml-core" rVer="1"><am3d:blip r:embed="${posterRid}"/></am3d:raster>`
		: '';
	return [
		'<am3d:camera><am3d:pos x="0" y="0" z="66519324"/><am3d:up dx="0" dy="36000000" dz="0"/>',
		'<am3d:lookAt x="0" y="0" z="0"/><am3d:perspective fov="2700000"/></am3d:camera>',
		'<am3d:trans><am3d:meterPerModelUnit n="1000000" d="1000000"/><am3d:preTrans dx="0" dy="0" dz="0"/>',
		'<am3d:scale><am3d:sx n="1000000" d="1000000"/><am3d:sy n="1000000" d="1000000"/>',
		'<am3d:sz n="1000000" d="1000000"/></am3d:scale><am3d:rot/><am3d:postTrans dx="0" dy="0" dz="0"/>',
		'</am3d:trans>',
		raster,
		'<am3d:objViewport viewportSz="3888064"/>',
		'<am3d:ambientLight><am3d:clr><a:scrgbClr r="50000" g="50000" b="50000"/></am3d:clr>',
		'<am3d:illuminance n="500000" d="1000000"/></am3d:ambientLight>',
	].join('');
}
