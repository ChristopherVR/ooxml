import { NS, elements, first, parseXml, type XmlDocument, type XmlElement } from '../xml/index';
import { CAMERA_PRESETS, LIGHT_DIRECTIONS, LIGHT_RIGS } from './scene3d-presets';
import {
	parseDrawingTextScene3D,
	parseDrawingTextShape3D,
	parseDrawingFlatText3D,
} from './read-text-3d';
import { patchDrawingColorIn } from './patch-color';
import type {
	DrawingScene3D,
	DrawingTextShape3D,
	DrawingTextScene3D,
	DrawingFlatText3D,
} from './text-3d-model';
import { scaleVectorToIntegers } from './scene3d-coordinate';
const bevels = new Set([
	'circle',
	'relaxedInset',
	'cross',
	'coolSlant',
	'angle',
	'softRound',
	'convex',
	'slope',
	'divot',
	'riblet',
	'hardEdge',
	'artDeco',
]);
const materials = new Set([
	'legacyMatte',
	'legacyPlastic',
	'legacyMetal',
	'legacyWireframe',
	'matte',
	'plastic',
	'metal',
	'warmMatte',
	'translucentPowder',
	'powder',
	'dkEdge',
	'softEdge',
	'clear',
	'flat',
	'softmetal',
]);
function remove(node: XmlElement | undefined): void {
	if (node) node.parentNode?.removeChild(node);
}
function semantics(value: object): string {
	return JSON.stringify(
		Object.fromEntries(Object.entries(value).filter(([key]) => key !== 'sourceXml')),
	);
}
function source(doc: XmlDocument, name: string, xml?: string): XmlElement {
	if (!xml) return doc.createElementNS(NS.a, `a:${name}`);
	const element = parseXml(xml).documentElement;
	if (element.namespaceURI !== NS.a || element.localName !== name)
		throw new Error(`Invalid ${name} source`);
	return doc.importNode(element, true) as XmlElement;
}
function child(node: XmlElement, name: string): XmlElement {
	const existing = first(node, name, NS.a);
	if (existing) return existing;
	const result = node.ownerDocument.createElementNS(NS.a, `a:${name}`);
	const order =
		node.localName === 'scene3d'
			? ['camera', 'lightRig', 'backdrop', 'extLst']
			: node.localName === 'sp3d'
				? ['bevelT', 'bevelB', 'extrusionClr', 'contourClr', 'extLst']
				: node.localName === 'backdrop'
					? ['anchor', 'norm', 'up', 'extLst']
					: ['rot', 'extLst'];
	const rank = order.indexOf(name);
	node.insertBefore(
		result,
		elements(node).find(
			(item) => item.namespaceURI === NS.a && order.indexOf(item.localName) > rank,
		) ?? null,
	);
	return result;
}
function number(
	node: XmlElement,
	name: string,
	value: number | undefined,
	before: number | undefined,
	scale = 1,
	min = -Infinity,
	max = Infinity,
): void {
	if (value === before) return;
	if (value === undefined) {
		node.removeAttribute(name);
		return;
	}
	const encoded = Math.round(value * scale);
	if (!Number.isFinite(value) || !Number.isSafeInteger(encoded) || value < min || value > max)
		throw new RangeError(`Invalid 3D ${name}`);
	node.setAttribute(name, String(encoded));
}
function token(
	node: XmlElement,
	name: string,
	value: string | undefined,
	before: string | undefined,
	allowed: Set<string>,
): void {
	if (value === before) return;
	if (value === undefined) node.removeAttribute(name);
	else {
		if (!allowed.has(value)) throw new RangeError(`Invalid 3D ${name}`);
		node.setAttribute(name, value);
	}
}
function rotation(
	node: XmlElement,
	value: DrawingScene3D,
	before: DrawingScene3D,
	prefix: 'camera' | 'lightRig',
): void {
	const axes = [
		['lat', 'RotX'],
		['lon', 'RotY'],
		['rev', 'RotZ'],
	] as const;
	if (!axes.some(([, suffix]) => value[`${prefix}${suffix}`] !== before[`${prefix}${suffix}`]))
		return;
	const rot = child(node, 'rot');
	for (const [attribute, suffix] of axes)
		number(rot, attribute, value[`${prefix}${suffix}`] ?? 0, before[`${prefix}${suffix}`]);
}
export function writeDrawingTextScene3D(doc: XmlDocument, value: DrawingTextScene3D): XmlElement {
	const node = source(doc, 'scene3d', value.sourceXml);
	const before = value.sourceXml ? parseDrawingTextScene3D(node) : {};
	if (value.sourceXml && semantics(value) === semantics(before)) return node;
	if (!value.sourceXml && (!value.cameraPreset || !value.lightRigType || !value.lightRigDirection))
		throw new Error('A new scene needs camera and light rig presets');
	const camera = child(node, 'camera');
	const light = child(node, 'lightRig');
	token(camera, 'prst', value.cameraPreset, before.cameraPreset, CAMERA_PRESETS);
	token(light, 'rig', value.lightRigType, before.lightRigType, LIGHT_RIGS);
	token(light, 'dir', value.lightRigDirection, before.lightRigDirection, LIGHT_DIRECTIONS);
	number(camera, 'fov', value.cameraFieldOfView, before.cameraFieldOfView, 1, 0, 10800000);
	number(camera, 'zoom', value.cameraZoom, before.cameraZoom, 100000, Number.MIN_VALUE);
	rotation(camera, value, before, 'camera');
	rotation(light, value, before, 'lightRig');
	if (value.hasBackdrop === false && before.hasBackdrop) remove(first(node, 'backdrop', NS.a));
	else if (value.hasBackdrop) {
		if (
			!first(node, 'backdrop', NS.a) &&
			[
				'backdropAnchorX',
				'backdropAnchorY',
				'backdropAnchorZ',
				'backdropNormalX',
				'backdropNormalY',
				'backdropNormalZ',
				'backdropUpX',
				'backdropUpY',
				'backdropUpZ',
			].some((key) => value[key as keyof DrawingScene3D] === undefined)
		)
			throw new Error('A new backdrop needs anchor, normal and up vectors');
		const backdrop = child(node, 'backdrop');
		for (const [name, prefix, attr] of [
			['anchor', 'backdropAnchor', ''],
			['norm', 'backdropNormal', 'd'],
			['up', 'backdropUp', 'd'],
		] as const) {
			const values = [value[`${prefix}X`], value[`${prefix}Y`], value[`${prefix}Z`]];
			if (
				!values.some((v, index) => v !== before[`${prefix}${(['X', 'Y', 'Z'] as const)[index]!}`])
			)
				continue;
			if (values.some((v) => v === undefined || !Number.isFinite(v)))
				throw new Error('A backdrop point needs three finite coordinates');
			const point = child(backdrop, name);
			const vector =
				attr === 'd'
					? scaleVectorToIntegers(values[0]!, values[1]!, values[2]!)
					: { x: values[0]!, y: values[1]!, z: values[2]! };
			for (const [axis, suffix] of [
				['x', 'X'],
				['y', 'Y'],
				['z', 'Z'],
			] as const)
				number(point, `${attr}${axis}`, vector[axis], before[`${prefix}${suffix}`]);
		}
	}
	return node;
}
export function writeDrawingTextShape3D(doc: XmlDocument, value: DrawingTextShape3D): XmlElement {
	const node = source(doc, 'sp3d', value.sourceXml);
	const before = value.sourceXml ? parseDrawingTextShape3D(node) : {};
	if (value.sourceXml && semantics(value) === semantics(before)) return node;
	number(node, 'z', value.z, before.z);
	number(node, 'extrusionH', value.extrusionHeight, before.extrusionHeight, 1, 0);
	number(node, 'contourW', value.contourWidth, before.contourWidth, 1, 0);
	token(node, 'prstMaterial', value.presetMaterial, before.presetMaterial, materials);
	for (const [name, prefix] of [
		['bevelT', 'bevelTop'],
		['bevelB', 'bevelBottom'],
	] as const) {
		if (
			value[`${prefix}Type`] === before[`${prefix}Type`] &&
			value[`${prefix}Width`] === before[`${prefix}Width`] &&
			value[`${prefix}Height`] === before[`${prefix}Height`]
		)
			continue;
		if (!value[`${prefix}Type`] || value[`${prefix}Type`] === 'none') {
			remove(first(node, name, NS.a));
			continue;
		}
		const bevel = child(node, name);
		token(bevel, 'prst', value[`${prefix}Type`], before[`${prefix}Type`], bevels);
		number(bevel, 'w', value[`${prefix}Width`], before[`${prefix}Width`], 1, 0);
		number(bevel, 'h', value[`${prefix}Height`], before[`${prefix}Height`], 1, 0);
	}
	for (const [key, name] of [
		['extrusionColor', 'extrusionClr'],
		['contourColor', 'contourClr'],
	] as const) {
		if (JSON.stringify(value[key]) === JSON.stringify(before[key])) continue;
		if (!value[key]) remove(first(node, name, NS.a));
		else patchDrawingColorIn(child(node, name), value[key]!);
	}
	return node;
}
export function writeDrawingFlatText3D(doc: XmlDocument, value: DrawingFlatText3D): XmlElement {
	const node = source(doc, 'flatTx', value.sourceXml);
	number(node, 'z', value.z, value.sourceXml ? parseDrawingFlatText3D(node).z : undefined);
	return node;
}
