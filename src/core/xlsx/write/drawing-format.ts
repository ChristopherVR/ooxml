import { drawingEffectXml, drawingEffectsXml } from '../../drawingml/drawing-effects';
import type { DrawingEffectList, DrawingOuterShadow } from '../../drawingml/effect-model';
import type { DrawingFill, DrawingLine } from '../../drawingml/types';
import { drawingFillXml } from '../../drawingml/write-fill';
import { drawingLineXml } from '../../drawingml/write-line';
import { drawingShapeStyleRefsXml, type CompleteStyleRefs } from '../../drawingml/shape-style-refs';
import type { DrawingTextScene3D, DrawingTextShape3D } from '../../drawingml/text-3d-model';
import { writeDrawingTextScene3D, writeDrawingTextShape3D } from '../../drawingml/write-text-3d';
import { NS, buildXml, elements, first, parseXml, type XmlElement } from '../../xml/index';

import { escapeAttr } from './xml-out';

const geometry = ['prstGeom', 'custGeom'];
const fills = ['noFill', 'solidFill', 'gradFill', 'pattFill', 'blipFill', 'grpFill'];
const effects = ['effectLst', 'effectDag'];
const order = ['xfrm', ...geometry, ...fills, 'ln', ...effects, 'scene3d', 'sp3d', 'extLst'];
const effectOrder = [
	'blur',
	'fillOverlay',
	'glow',
	'innerShdw',
	'outerShdw',
	'prstShdw',
	'reflection',
	'softEdge',
];

export interface DrawingFrameFormat {
	geometry?: string;
	adjustments?: Readonly<Record<string, number>>;
	fill?: DrawingFill | null;
	line?: DrawingLine | null;
	effects?: DrawingEffectList | null;
	shadow?: DrawingOuterShadow | null;
	scene3d?: DrawingTextScene3D | null;
	shape3d?: DrawingTextShape3D | null;
	style?: CompleteStyleRefs;
}

function replace(
	parent: XmlElement,
	names: readonly string[],
	xml: string | undefined,
	schemaOrder = order,
): void {
	const doc = parent.ownerDocument;
	const children = elements(parent).filter((child) => child.namespaceURI === NS.a);
	const old = children.filter((child) => names.includes(child.localName));
	if (xml) {
		const next = elements(parseXml(`<root xmlns:a="${NS.a}">${xml}</root>`).documentElement)[0];
		if (next) {
			const rank = schemaOrder.indexOf(next.localName);
			const before =
				old[0] ?? children.find((child) => schemaOrder.indexOf(child.localName) > rank);
			parent.insertBefore(doc.importNode(next, true), before ?? null);
		}
	}
	for (const child of old) parent.removeChild(child);
}

export function applyDrawingFrameFormat(node: XmlElement, format: DrawingFrameFormat): void {
	const props = first(node, 'spPr', NS.xdr);
	if (!props) return;
	if (format.geometry !== undefined) {
		const guides = Object.entries(format.adjustments ?? {})
			.map(([name, value]) => `<a:gd name="${escapeAttr(name)}" fmla="val ${value}"/>`)
			.join('');
		replace(
			props,
			geometry,
			`<a:prstGeom prst="${escapeAttr(format.geometry)}"><a:avLst>${guides}</a:avLst></a:prstGeom>`,
		);
	}
	if (format.fill !== undefined)
		replace(props, fills, format.fill ? drawingFillXml(format.fill) : undefined);
	if (format.line !== undefined)
		replace(props, ['ln'], format.line ? drawingLineXml(format.line) : undefined);
	if (format.effects !== undefined)
		replace(props, effects, format.effects ? drawingEffectsXml(format.effects) : undefined);
	const doc = node.ownerDocument;
	if (format.scene3d !== undefined)
		replace(
			props,
			['scene3d'],
			format.scene3d ? buildXml(writeDrawingTextScene3D(doc, format.scene3d)) : undefined,
		);
	if (format.shape3d !== undefined)
		replace(
			props,
			['sp3d'],
			format.shape3d ? buildXml(writeDrawingTextShape3D(doc, format.shape3d)) : undefined,
		);
	if (format.style) {
		const old = first(node, 'style', NS.xdr);
		const refs = drawingShapeStyleRefsXml(format.style);
		const next = parseXml(
			`<xdr:style xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}">${refs}</xdr:style>`,
		).documentElement;
		const style = doc.importNode(next, true);
		if (old) node.replaceChild(style, old);
		else node.insertBefore(style, props.nextSibling);
	}
	if (format.shadow === undefined) return;
	let list = first(props, 'effectLst', NS.a);
	if (!list) {
		replace(props, effects, '<a:effectLst/>');
		list = first(props, 'effectLst', NS.a);
	}
	if (list)
		replace(
			list,
			['outerShdw'],
			format.shadow ? drawingEffectXml(format.shadow) : undefined,
			effectOrder,
		);
}

export function drawingIdentity(node: XmlElement): XmlElement | undefined {
	const nv = elements(node).find(
		(child) => child.namespaceURI === NS.xdr && child.localName.startsWith('nv'),
	);
	return first(nv, 'cNvPr', NS.xdr);
}

export function applyDrawingDescription(node: XmlElement, description: string): void {
	const identity = drawingIdentity(node);
	if (description) identity?.setAttribute('descr', description);
	else identity?.removeAttribute('descr');
}

export function applyDrawingExtent(node: XmlElement, width: number, height: number): void {
	const props = first(node, node.localName === 'grpSp' ? 'grpSpPr' : 'spPr', NS.xdr);
	const ext = first(first(props, 'xfrm', NS.a), 'ext', NS.a);
	ext?.setAttribute('cx', String(width));
	ext?.setAttribute('cy', String(height));
}
