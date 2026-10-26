import { NS, elements, first, parseXml, type XmlElement } from '../../xml/index';
import type { DrawingColor } from '../../drawingml/types';
import { patchDrawingColorIn } from '../../drawingml/patch-color';
import { parseDrawingColorIn } from '../../drawingml/drawing-color';
import type { ChartSpace, ChartView3D } from '../model';
import type { ChartShapeProperties } from '../model-series';
import { CHART_ROOT_BINDINGS, innerXml } from '../xml-fragment';
import { editChartSpace } from './draft';
import type { ChartFormatTarget } from './target';
import { chartTargetDraftShapes, chartTargetShape } from './target-paint';

export interface ChartBevelEdit {
	preset: string;
	widthEmu: number;
	heightEmu: number;
}
export interface ChartThreeDEdit {
	bevelTop?: ChartBevelEdit | null;
	bevelBottom?: ChartBevelEdit | null;
	depthEmu?: number | null;
	depthColor?: DrawingColor | null;
	contourWidthEmu?: number | null;
	contourColor?: DrawingColor | null;
	material?: string | null;
	lighting?: string | null;
}

function child(parent: XmlElement, local: string): XmlElement {
	const existing = first(parent, local, NS.a);
	if (existing) return existing;
	const node = parent.ownerDocument.createElementNS(NS.a, `a:${local}`);
	parent.insertBefore(node, first(parent, 'extLst', NS.a) ?? null);
	return node;
}

/** Applies a DOM-only shape edit to the existing source; the ordinary shape writer retains it. */
export function editChartShapeXml(
	shape: ChartShapeProperties,
	edit: (root: XmlElement) => void,
): void {
	const declarations = [...CHART_ROOT_BINDINGS]
		.map(([prefix, uri]) => ` xmlns:${prefix}="${uri}"`)
		.join('');
	const root = parseXml(`<c:spPr${declarations}>${shape.sourceXml ?? ''}</c:spPr>`).documentElement;
	edit(root);
	shape.sourceXml = innerXml(root);
}

export function chartTargetThreeD(space: ChartSpace, target: ChartFormatTarget): ChartThreeDEdit {
	const shape = chartTargetShape(space, target);
	const result: ChartThreeDEdit = {};
	if (!shape?.sourceXml) return result;
	editChartShapeXml({ ...shape }, (root) => {
		const sp3d = first(root, 'sp3d', NS.a);
		if (sp3d) {
			for (const [key, tag] of [
				['bevelTop', 'bevelT'],
				['bevelBottom', 'bevelB'],
			] as const) {
				const node = first(sp3d, tag, NS.a);
				if (!node) continue;
				const widthEmu = Number(node.getAttribute('w') ?? 0);
				const heightEmu = Number(node.getAttribute('h') ?? 0);
				if ([widthEmu, heightEmu].every(Number.isFinite))
					result[key] = { preset: node.getAttribute('prst') ?? 'circle', widthEmu, heightEmu };
			}
			for (const [key, attr] of [
				['depthEmu', 'extrusionH'],
				['contourWidthEmu', 'contourW'],
			] as const) {
				const value = sp3d.getAttribute(attr);
				if (value !== null && Number.isFinite(Number(value))) result[key] = Number(value);
			}
			for (const [key, tag] of [
				['depthColor', 'extrusionClr'],
				['contourColor', 'contourClr'],
			] as const) {
				const color = parseDrawingColorIn(first(sp3d, tag, NS.a));
				if (color) result[key] = color;
			}
			const material = sp3d.getAttribute('prstMaterial');
			if (material !== null) result.material = material;
		}
		const rig = first(first(root, 'scene3d', NS.a), 'lightRig', NS.a)?.getAttribute('rig');
		if (rig !== null && rig !== undefined) result.lighting = rig;
	});
	return result;
}

function remove(parent: XmlElement, local: string): void {
	const old = first(parent, local, NS.a);
	if (old) parent.removeChild(old);
}

/** Shape bevel/extrusion controls, retaining unmodelled 3-D/camera and extension XML. */
export function setChartTargetThreeD(
	space: ChartSpace,
	target: ChartFormatTarget,
	edit: ChartThreeDEdit,
): ChartSpace {
	return editChartSpace(space, (draft) => {
		for (const shape of chartTargetDraftShapes(draft, target))
			editChartShapeXml(shape, (root) => {
				const sp3d = child(root, 'sp3d');
				for (const [key, tag] of [
					['bevelTop', 'bevelT'],
					['bevelBottom', 'bevelB'],
				] as const) {
					const bevel = edit[key];
					if (bevel === undefined) continue;
					if (bevel === null) {
						remove(sp3d, tag);
						continue;
					}
					if (
						![bevel.widthEmu, bevel.heightEmu].every(
							(value) => Number.isFinite(value) && value >= 0,
						) ||
						!bevel.preset
					)
						throw new RangeError('Invalid chart bevel');
					const node = child(sp3d, tag);
					node.setAttribute('prst', bevel.preset);
					node.setAttribute('w', String(Math.round(bevel.widthEmu)));
					node.setAttribute('h', String(Math.round(bevel.heightEmu)));
				}
				for (const [key, attribute] of [
					['depthEmu', 'extrusionH'],
					['contourWidthEmu', 'contourW'],
				] as const) {
					const value = edit[key];
					if (value === undefined) continue;
					if (value === null) sp3d.removeAttribute(attribute);
					else {
						if (!Number.isFinite(value) || value < 0) throw new RangeError('Invalid chart 3D size');
						sp3d.setAttribute(attribute, String(Math.round(value)));
					}
				}
				for (const [key, tag] of [
					['depthColor', 'extrusionClr'],
					['contourColor', 'contourClr'],
				] as const) {
					const color = edit[key];
					if (color === undefined) continue;
					if (color === null) remove(sp3d, tag);
					else patchDrawingColorIn(child(sp3d, tag), color);
				}
				if (edit.material !== undefined) {
					if (edit.material === null) sp3d.removeAttribute('prstMaterial');
					else sp3d.setAttribute('prstMaterial', edit.material);
				}
				if (edit.lighting !== undefined) {
					{
						const scene =
							first(root, 'scene3d', NS.a) ?? root.ownerDocument.createElementNS(NS.a, 'a:scene3d');
						if (!scene.parentNode) root.insertBefore(scene, sp3d);
						if (!first(scene, 'camera', NS.a))
							child(scene, 'camera').setAttribute('prst', 'orthographicFront');
						const rig = child(scene, 'lightRig');
						rig.setAttribute('rig', edit.lighting ?? 'threePt');
						if (!rig.hasAttribute('dir')) rig.setAttribute('dir', 't');
					}
				}
				// CT_Shape3D's fixed child order; unknown nodes remain alongside their source data.
				const names = ['bevelT', 'bevelB', 'extrusionClr', 'contourClr', 'extLst'];
				const known = elements(sp3d).filter(
					(node) => node.namespaceURI === NS.a && names.includes(node.localName),
				);
				for (const node of known.sort(
					(a, b) => names.indexOf(a.localName) - names.indexOf(b.localName),
				))
					sp3d.appendChild(node);
			});
	});
}

/** The chart's 3-D viewing camera. Two-dimensional families cannot author a 3-D view. */
export function setChartView3D(space: ChartSpace, edit: Partial<ChartView3D>): ChartSpace {
	if (!space.plotArea.groups.some((group) => group.is3D))
		throw new RangeError('The chart has no 3D plot group');
	const bounds: Record<string, readonly [number, number]> = {
		rotX: [-90, 90],
		rotY: [0, 360],
		depthPercent: [20, 2000],
		heightPercent: [5, 500],
		perspective: [0, 240],
	};
	for (const [key, value] of Object.entries(edit)) {
		const range = bounds[key];
		if (
			range &&
			(typeof value !== 'number' || !Number.isFinite(value) || value < range[0] || value > range[1])
		)
			throw new RangeError(`Invalid chart 3D ${key}`);
	}
	return editChartSpace(space, (draft) => {
		draft.view3D = { ...draft.view3D, ...edit };
	});
}
