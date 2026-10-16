/**
 * Writes and reads the 3-D Format section of a chart element: `a:sp3d`
 * (top and bottom bevel, depth, contour, material) and the lighting rig of
 * `a:scene3d`, inside the element's `c:spPr`.
 *
 * @module utils/chart-element-format-3d-xml
 */
import type {
	PptxChartBevel,
	PptxChartBevelType,
	PptxChartElementThreeD,
	XmlObject,
} from '../types';
import type { LocalName } from './chart-element-format-xml-util';
import {
	EMU_PER_POINT,
	asObject,
	chartElementColorXml,
	childOf,
	keyOf,
	pt,
	removeChildren,
	setOrdered,
} from './chart-element-format-xml-util';

const SP3D_ORDER = ['bevelT', 'bevelB', 'extrusionClr', 'contourClr', 'extLst'];
const BEVELS = new Set<string>([
	'relaxedInset',
	'circle',
	'slope',
	'cross',
	'angle',
	'softRound',
	'convex',
	'coolSlant',
	'divot',
	'riblet',
	'hardEdge',
	'artDeco',
]);

function bevelXml(bevel: PptxChartBevel): XmlObject {
	return { '@_w': pt(bevel.width), '@_h': pt(bevel.height), '@_prst': bevel.type };
}

function setBevel(
	sp3d: XmlObject,
	name: 'bevelT' | 'bevelB',
	bevel: PptxChartBevel | null | undefined,
	localName: LocalName,
): void {
	if (bevel === undefined) {
		return;
	}
	removeChildren(sp3d, [name], localName);
	if (bevel) {
		setOrdered(sp3d, name, bevelXml(bevel), SP3D_ORDER, localName);
	}
}

function setSized(
	sp3d: XmlObject,
	attr: '@_extrusionH' | '@_contourW',
	colorName: 'extrusionClr' | 'contourClr',
	value: { size: number; color?: string } | null | undefined,
	localName: LocalName,
): void {
	if (value === undefined) {
		return;
	}
	removeChildren(sp3d, [colorName], localName);
	if (!value || value.size <= 0) {
		delete sp3d[attr];
		return;
	}
	sp3d[attr] = pt(value.size);
	if (value.color) {
		setOrdered(sp3d, colorName, chartElementColorXml(value.color), SP3D_ORDER, localName);
	}
}

/**
 * Apply `threeD` to `spPr` (in place), creating `a:sp3d` / `a:scene3d` at
 * their schema positions through `setSpPrChild`.
 */
export function applyChartThreeD(
	spPr: XmlObject,
	threeD: PptxChartElementThreeD,
	localName: LocalName,
	setSpPrChild: (name: string, value: XmlObject | undefined) => void,
): void {
	const sp3dKey = keyOf(spPr, 'sp3d', localName);
	const sp3d: XmlObject = { ...(asObject(sp3dKey ? spPr[sp3dKey] : undefined) ?? {}) };
	setBevel(sp3d, 'bevelT', threeD.bevelTop, localName);
	setBevel(sp3d, 'bevelB', threeD.bevelBottom, localName);
	setSized(sp3d, '@_extrusionH', 'extrusionClr', threeD.depth, localName);
	setSized(sp3d, '@_contourW', 'contourClr', threeD.contour, localName);
	if (threeD.material !== undefined) {
		if (threeD.material) sp3d['@_prstMaterial'] = threeD.material;
		else delete sp3d['@_prstMaterial'];
	}
	setSpPrChild('sp3d', Object.keys(sp3d).length > 0 ? sp3d : undefined);
	if (threeD.lighting !== undefined) {
		const sceneKey = keyOf(spPr, 'scene3d', localName);
		if (!threeD.lighting) {
			setSpPrChild('scene3d', undefined);
			return;
		}
		const scene: XmlObject = { ...(asObject(sceneKey ? spPr[sceneKey] : undefined) ?? {}) };
		const camera = keyOf(scene, 'camera', localName);
		const rigKey = keyOf(scene, 'lightRig', localName);
		const rig: XmlObject = {
			'@_dir': 't',
			...(asObject(rigKey ? scene[rigKey] : undefined) ?? {}),
			'@_rig': threeD.lighting,
		};
		const next: XmlObject = {};
		next[camera ?? 'a:camera'] = camera ? scene[camera] : { '@_prst': 'orthographicFront' };
		next[rigKey ?? 'a:lightRig'] = rig;
		for (const [key, value] of Object.entries(scene)) {
			if (key !== camera && key !== rigKey) next[key] = value;
		}
		setSpPrChild('scene3d', next);
	}
}

const points = (emu: unknown): number => {
	const n = Number(emu);
	return Number.isFinite(n) ? Math.round((n / EMU_PER_POINT) * 100) / 100 : 0;
};

function readBevel(node: XmlObject | undefined): PptxChartBevel | undefined {
	if (!node) {
		return undefined;
	}
	const prst = String(node['@_prst'] ?? 'circle');
	return {
		type: (BEVELS.has(prst) ? prst : 'circle') as PptxChartBevelType,
		// ECMA-376 defaults: 76200 EMU (6pt) wide and high.
		width: node['@_w'] === undefined ? 6 : points(node['@_w']),
		height: node['@_h'] === undefined ? 6 : points(node['@_h']),
	};
}

/** The 3-D Format a `c:spPr` carries, or `undefined` when none. */
export function readChartThreeD(
	spPr: XmlObject,
	localName: LocalName,
	resolve: (fill: XmlObject) => string | undefined,
): PptxChartElementThreeD | undefined {
	const out: PptxChartElementThreeD = {};
	const sp3d = childOf(spPr, 'sp3d', localName);
	if (sp3d) {
		const top = readBevel(childOf(sp3d, 'bevelT', localName));
		const bottom = readBevel(childOf(sp3d, 'bevelB', localName));
		if (top) out.bevelTop = top;
		if (bottom) out.bevelBottom = bottom;
		const depth = points(sp3d['@_extrusionH'] ?? 0);
		if (depth > 0) {
			const clr = childOf(sp3d, 'extrusionClr', localName);
			const color = clr ? resolve(clr) : undefined;
			out.depth = { size: depth, ...(color ? { color } : {}) };
		}
		const contour = points(sp3d['@_contourW'] ?? 0);
		if (contour > 0) {
			const clr = childOf(sp3d, 'contourClr', localName);
			const color = clr ? resolve(clr) : undefined;
			out.contour = { size: contour, ...(color ? { color } : {}) };
		}
		if (typeof sp3d['@_prstMaterial'] === 'string') out.material = sp3d['@_prstMaterial'];
	}
	const rig = childOf(childOf(spPr, 'scene3d', localName), 'lightRig', localName)?.['@_rig'];
	if (typeof rig === 'string') out.lighting = rig;
	return Object.keys(out).length > 0 ? out : undefined;
}
