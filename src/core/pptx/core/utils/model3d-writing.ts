import { MODEL3D_URI } from '../../../drawingml/model3d';
import type { XmlObject } from '../types';

export { defaultModel3DSceneXml, MODEL3D_URI, MODEL3D_REL } from '../../../drawingml/model3d';

export function readModel3DGraphic(frame: XmlObject | undefined): XmlObject | undefined {
	const graphic = frame?.['a:graphic'] as XmlObject | undefined;
	const data = graphic?.['a:graphicData'] as XmlObject | undefined;
	return data?.['@_uri'] === MODEL3D_URI
		? (data['am3d:model3d'] as XmlObject | undefined)
		: undefined;
}
