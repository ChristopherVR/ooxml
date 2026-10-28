import { readScene3DValues } from '../../../drawingml/read-3d-values';
import type { TextStyle, XmlObject } from '../types';
import { cloneXmlObject } from './clone-utils';
import { roundCoordinate, scaleVectorToIntegers } from './scene3d-coordinate';
import { CAMERA_PRESETS, LIGHT_DIRECTIONS, LIGHT_RIGS } from './text-body-scene3d-presets';

const localName = (key: string): string => key.replace(/^@_/, '').split(':').pop() ?? key;
const childKey = (node: XmlObject, local: string): string | undefined =>
	Object.keys(node).find((key) => !key.startsWith('@_') && localName(key) === local);
const child = (node: XmlObject | undefined, local: string): XmlObject | undefined => {
	if (!node) {
		return undefined;
	}
	const key = childKey(node, local);
	return key ? (node[key] as XmlObject | undefined) : undefined;
};
const enumToken = (value: unknown, allowed: Set<string>): string | undefined => {
	const token = String(value ?? '').trim();
	return allowed.has(token) ? token : undefined;
};
export function parseTextBodyScene3d(bodyPr: XmlObject, style: TextStyle): void {
	const scene = child(bodyPr, 'scene3d');
	if (!scene) {
		return;
	}
	style.textBodyScene3dXml = cloneXmlObject(scene);
	const typed = readScene3DValues(scene, { child, attribute: (node, name) => node?.[`@_${name}`] });
	style.textBodyScene3d = {
		cameraPreset: undefined,
		lightRigType: undefined,
		lightRigDirection: undefined,
		...typed,
	};
}

const setAttrs = (node: XmlObject, attrs: Record<string, number | string | undefined>): void => {
	for (const [key, value] of Object.entries(attrs)) {
		if (value !== undefined) {
			node[`@_${key}`] = String(value);
		}
	}
};
const setChild = (node: XmlObject, local: string, value: XmlObject): void => {
	const key = childKey(node, local) ?? `a:${local}`;
	node[key] = value;
};
const ordered = (node: XmlObject, order: readonly string[]): XmlObject => {
	const result: XmlObject = {};
	for (const local of order) {
		const key = childKey(node, local);
		if (key) {
			result[key] = node[key];
		}
	}
	for (const key of Object.keys(node)) {
		if (!Object.hasOwn(result, key)) {
			result[key] = node[key];
		}
	}
	return result;
};
const rotationXml = (x?: number, y?: number, z?: number): XmlObject | undefined => {
	if (
		[x, y, z].some(
			(value) =>
				value !== undefined && (!Number.isSafeInteger(value) || value < 0 || value >= 21600000),
		)
	)
		throw new RangeError('Invalid 3D rotation angle');
	if (![x, y, z].some(Number.isInteger)) {
		return undefined;
	}
	const rot: XmlObject = {};
	setAttrs(rot, { lat: x, lon: y, rev: z });
	return rot;
};
/**
 * `CT_Point3D` (`a:anchor`): a position, `ST_Coordinate` (integer). Each
 * component is rounded independently; unlike a direction vector, there is no
 * ratio between x/y/z to preserve.
 */
const anchorXml = (x?: number, y?: number, z?: number): XmlObject | undefined => {
	if (x === undefined || y === undefined || z === undefined) {
		return undefined;
	}
	const point: XmlObject = {};
	setAttrs(point, { x: roundCoordinate(x), y: roundCoordinate(y), z: roundCoordinate(z) });
	return point;
};

/**
 * `CT_Vector3D` (`a:norm`/`a:up`): a direction, `ST_Coordinate` (integer)
 * components where only the ratio between dx/dy/dz matters. Scaled via
 * {@link scaleVectorToIntegers} rather than the strict "every component must
 * already be an integer, else drop the whole node" the old `pointXml` did,
 * which silently discarded an authored backdrop the moment any one component
 * (e.g. a normalised float direction) was fractional.
 */
const vectorXml = (x?: number, y?: number, z?: number): XmlObject | undefined => {
	if (x === undefined || y === undefined || z === undefined) {
		return undefined;
	}
	const { x: dx, y: dy, z: dz } = scaleVectorToIntegers(x, y, z);
	const vector: XmlObject = {};
	setAttrs(vector, { dx, dy, dz });
	return vector;
};

export function applyTextBodyScene3d(bodyPr: XmlObject, style: TextStyle | undefined): void {
	const existingKey = childKey(bodyPr, 'scene3d');
	const raw =
		cloneXmlObject(style?.textBodyScene3dXml) ??
		(existingKey ? cloneXmlObject(bodyPr[existingKey] as XmlObject) : undefined);
	const typed = style?.textBodyScene3d;
	if (!typed) {
		if (raw) {
			bodyPr[existingKey ?? 'a:scene3d'] = raw;
		}
		return;
	}
	const cameraPreset = enumToken(typed.cameraPreset, CAMERA_PRESETS);
	const lightRig = enumToken(typed.lightRigType, LIGHT_RIGS);
	const lightDir = enumToken(typed.lightRigDirection, LIGHT_DIRECTIONS);
	if ((!cameraPreset || !lightRig || !lightDir) && !raw) {
		return;
	}
	const scene = raw ?? {};
	const camera = cloneXmlObject(child(scene, 'camera')) ?? {};
	if (cameraPreset) {
		camera['@_prst'] = cameraPreset;
	}
	if (
		typed.cameraFieldOfView !== undefined &&
		Number.isInteger(typed.cameraFieldOfView) &&
		typed.cameraFieldOfView >= 0 &&
		typed.cameraFieldOfView <= 10800000
	) {
		camera['@_fov'] = String(typed.cameraFieldOfView);
	}
	if (typed.cameraZoom !== undefined && Number.isFinite(typed.cameraZoom) && typed.cameraZoom > 0) {
		camera['@_zoom'] = String(Math.round(typed.cameraZoom * 100000));
	}
	const cameraRot = rotationXml(typed.cameraRotX, typed.cameraRotY, typed.cameraRotZ);
	if (cameraRot) {
		setChild(camera, 'rot', cameraRot);
	}
	setChild(scene, 'camera', ordered(camera, ['rot']));
	const light = cloneXmlObject(child(scene, 'lightRig')) ?? {};
	if (lightRig) {
		light['@_rig'] = lightRig;
	}
	if (lightDir) {
		light['@_dir'] = lightDir;
	}
	const lightRot = rotationXml(typed.lightRigRotX, typed.lightRigRotY, typed.lightRigRotZ);
	if (lightRot) {
		setChild(light, 'rot', lightRot);
	}
	setChild(scene, 'lightRig', ordered(light, ['rot']));
	if (typed.hasBackdrop) {
		const backdrop = cloneXmlObject(child(scene, 'backdrop')) ?? {};
		const anchor = anchorXml(typed.backdropAnchorX, typed.backdropAnchorY, typed.backdropAnchorZ);
		const norm = vectorXml(typed.backdropNormalX, typed.backdropNormalY, typed.backdropNormalZ);
		const up = vectorXml(typed.backdropUpX, typed.backdropUpY, typed.backdropUpZ);
		if (anchor) {
			setChild(backdrop, 'anchor', anchor);
		}
		if (norm) {
			setChild(backdrop, 'norm', norm);
		}
		if (up) {
			setChild(backdrop, 'up', up);
		}
		if (child(backdrop, 'anchor') && child(backdrop, 'norm') && child(backdrop, 'up')) {
			setChild(scene, 'backdrop', ordered(backdrop, ['anchor', 'norm', 'up', 'extLst']));
		}
	}
	bodyPr[existingKey ?? 'a:scene3d'] = ordered(scene, ['camera', 'lightRig', 'backdrop', 'extLst']);
}
