export const MODEL3D_URI = 'http://schemas.microsoft.com/office/drawing/2017/model3d';
export const MODEL3D_REL = 'http://schemas.microsoft.com/office/2017/06/relationships/model3d';
export const GLB_MIME = 'model/gltf-binary';
/** Content type used by native Office packages. */
export const GLB_PART_CONTENT_TYPE = 'model/gltf.binary';

/** Default camera and scene, using the public MS-ODRAWXML CT_Model3D child order. */
export function defaultModel3DSceneXml(posterRid?: string): string {
	const poster = posterRid
		?.replaceAll('&', '&amp;')
		.replaceAll('"', '&quot;')
		.replaceAll('<', '&lt;');
	const raster = poster
		? `<am3d:raster rName="ooxml-core" rVer="1"><am3d:blip r:embed="${poster}"/></am3d:raster>`
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
