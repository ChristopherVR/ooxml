import { drawingEffectXml, parseDrawingEffects } from '../../drawingml/drawing-effects';
import type { DrawingOuterShadow } from '../../drawingml/effect-model';
import type { DrawingFill, DrawingLine } from '../../drawingml/types';
import { drawingFillXml } from '../../drawingml/write-fill';
import { drawingLineXml } from '../../drawingml/write-line';
import { NS, elements, first, parseXml, type XmlElement } from '../../xml/index';

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
	fill?: DrawingFill;
	line?: DrawingLine;
	shadow?: DrawingOuterShadow | null;
}

export function readDrawingShadow(props: XmlElement | undefined): DrawingOuterShadow | undefined {
	return parseDrawingEffects(first(props, 'effectLst', NS.a))?.effects.find(
		(effect) => effect.kind === 'outerShadow',
	);
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
	if (format.geometry !== undefined)
		replace(props, geometry, `<a:prstGeom prst="${format.geometry}"><a:avLst/></a:prstGeom>`);
	if (format.fill) replace(props, fills, drawingFillXml(format.fill));
	if (format.line) replace(props, ['ln'], drawingLineXml(format.line));
	if (format.shadow === undefined) return;
	let list = first(props, 'effectLst', NS.a);
	if (!list && format.shadow) {
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
