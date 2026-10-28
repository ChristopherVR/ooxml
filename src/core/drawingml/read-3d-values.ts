import type { DrawingScene3D, DrawingText3DValues } from './text-3d-model';
import { CAMERA_PRESETS, LIGHT_DIRECTIONS, LIGHT_RIGS } from './scene3d-presets';
import { picturePercentage } from './picture-effects';
export interface ThreeDReader<T> {
	attribute(node: T | undefined, name: string): unknown;
	child(node: T | undefined, name: string): T | undefined;
}
const integer = (value: unknown): number | undefined => {
	if (value === undefined || value === null || String(value).trim() === '') return undefined;
	const n = Number(value);
	return Number.isSafeInteger(n) ? n : undefined;
};
export function readScene3DValues<T>(node: T, read: ThreeDReader<T>): DrawingScene3D {
	const out: DrawingScene3D = {};
	const camera = read.child(node, 'camera');
	const light = read.child(node, 'lightRig');
	for (const [key, source, attribute, choices] of [
		['cameraPreset', camera, 'prst', CAMERA_PRESETS],
		['lightRigType', light, 'rig', LIGHT_RIGS],
		['lightRigDirection', light, 'dir', LIGHT_DIRECTIONS],
	] as const) {
		const token = String(read.attribute(source, attribute) ?? '').trim();
		if (choices.has(token)) out[key] = token;
	}
	const fov = integer(read.attribute(camera, 'fov'));
	if (fov !== undefined && fov >= 0 && fov <= 10800000) out.cameraFieldOfView = fov;
	const zoom = picturePercentage(read.attribute(camera, 'zoom'));
	if (zoom !== undefined && zoom > 0) out.cameraZoom = zoom;
	for (const [prefix, source] of [
		['camera', camera],
		['lightRig', light],
	] as const) {
		const rot = read.child(source, 'rot');
		for (const [axis, key] of [
			['lat', 'RotX'],
			['lon', 'RotY'],
			['rev', 'RotZ'],
		] as const) {
			const value = integer(read.attribute(rot, axis));
			if (value !== undefined && value >= 0 && value < 21600000) out[`${prefix}${key}`] = value;
		}
	}
	const backdrop = read.child(node, 'backdrop');
	if (backdrop !== undefined) {
		out.hasBackdrop = true;
		for (const [name, prefix, attr] of [
			['anchor', 'backdropAnchor', ''],
			['norm', 'backdropNormal', 'd'],
			['up', 'backdropUp', 'd'],
		] as const) {
			const point = read.child(backdrop, name);
			for (const [axis, suffix] of [
				['x', 'X'],
				['y', 'Y'],
				['z', 'Z'],
			] as const) {
				const value = integer(read.attribute(point, `${attr}${axis}`));
				if (value !== undefined) out[`${prefix}${suffix}`] = value;
			}
		}
	}
	return out;
}
export function readText3DValues<T>(node: T, read: ThreeDReader<T>): DrawingText3DValues {
	const out: DrawingText3DValues = {};
	for (const [key, attribute] of [
		['extrusionHeight', 'extrusionH'],
		['contourWidth', 'contourW'],
		['z', 'z'],
	] as const) {
		const value = integer(read.attribute(node, attribute));
		if (value !== undefined) out[key] = value;
	}
	const material = String(read.attribute(node, 'prstMaterial') ?? '').trim();
	if (material) out.presetMaterial = material;
	for (const [name, prefix] of [
		['bevelT', 'bevelTop'],
		['bevelB', 'bevelBottom'],
	] as const) {
		const bevel = read.child(node, name);
		if (!bevel) continue;
		out[`${prefix}Type`] = String(read.attribute(bevel, 'prst') || 'circle').trim();
		for (const [attribute, suffix] of [
			['w', 'Width'],
			['h', 'Height'],
		] as const) {
			const value = integer(read.attribute(bevel, attribute));
			if (value !== undefined) out[`${prefix}${suffix}`] = value;
		}
	}
	return out;
}
