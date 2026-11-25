import { drawingFillXml } from '../../drawingml/write-fill';
import { drawingLineXml } from '../../drawingml/write-line';
import { NS, buildXml, elements, first, parseXml, type XmlElement } from '../../xml/index';
import type { ShapeNode } from '../model-shapes';
import { attrs, escapeText } from './xml-out';
import { wordArtShapeXml } from '../word-art';
import { equationShapeBranches, wrapEquationShape } from '../shape-equations';

const offset = (name: string, x: number, y: number) =>
	`<a:${name}${attrs({ x: Math.round(x), y: Math.round(y) })}/>`;
const extent = (name: string, width: number, height: number) =>
	`<a:${name}${attrs({ cx: Math.round(width), cy: Math.round(height) })}/>`;

function transform(shape: ShapeNode): string {
	const { frame, childFrame, rotation, flipH, flipV } = shape;
	return (
		`<a:xfrm${attrs({ rot: rotation ? Math.round(rotation * 60000) : undefined, flipH, flipV })}>` +
		offset('off', frame.x, frame.y) +
		extent('ext', frame.width, frame.height) +
		(childFrame
			? offset('chOff', childFrame.x, childFrame.y) +
				extent('chExt', childFrame.width, childFrame.height)
			: '') +
		'</a:xfrm>'
	);
}

function textBody(text: string): string {
	return (
		'<xdr:txBody><a:bodyPr/><a:lstStyle/>' +
		text
			.split(/\r?\n/)
			.map((line) => `<a:p><a:r><a:t>${escapeText(line)}</a:t></a:r></a:p>`)
			.join('') +
		'</xdr:txBody>'
	);
}

/** Native shape content. A new group assigns ids to children from the same drawing-wide counter. */
export function nativeShapeXml(shape: ShapeNode, nextId: () => number): string {
	const id = nextId();
	const nv = `<xdr:cNvPr${attrs({ id, name: shape.name ?? `Shape ${id}`, hidden: shape.hidden })}/>`;
	if (shape.sourceShapeXml) {
		const doc = parseXml(shape.sourceShapeXml);
		const node = doc.documentElement;
		const identity = doc.getElementsByTagNameNS(NS.xdr, 'cNvPr')[0];
		identity?.setAttribute('id', String(id));
		if (shape.hidden) identity?.setAttribute('hidden', '1');
		else identity?.removeAttribute('hidden');
		patchTransform(node, shape, true);
		return buildXml(node);
	}
	if (shape.wordArt) {
		const doc = parseXml(
			`<root xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}">${wordArtShapeXml(shape, id)}</root>`,
		);
		const node = elements(doc.documentElement)[0];
		if (node) {
			patchTransform(node, shape, true);
			if (shape.hidden) doc.getElementsByTagNameNS(NS.xdr, 'cNvPr')[0]?.setAttribute('hidden', '1');
			return buildXml(node);
		}
	}
	if (shape.kind === 'group')
		return `<xdr:grpSp><xdr:nvGrpSpPr>${nv}<xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr>${transform(shape)}</xdr:grpSpPr>${(shape.children ?? []).map((child) => nativeShapeXml(child, nextId)).join('')}</xdr:grpSp>`;
	const connector = shape.kind === 'connector';
	const tag = connector ? 'cxnSp' : 'sp';
	const nonVisual = connector ? 'CxnSpPr' : 'SpPr';
	const geometry = `<a:prstGeom${attrs({ prst: shape.geometry })}><a:avLst/></a:prstGeom>`;
	const fill = shape.fill ? (drawingFillXml(shape.fill) ?? '') : '';
	const line = shape.line ? drawingLineXml(shape.line) : '';
	const text = shape.textBox ? textBody(shape.text?.text ?? '') : '';
	const base = `<xdr:${tag}><xdr:nv${nonVisual}>${nv}<xdr:cNv${nonVisual}${shape.textBox ? ' txBox="1"' : ''}/></xdr:nv${nonVisual}><xdr:spPr>${transform(shape)}${geometry}${fill}${line}</xdr:spPr>`;
	const equation = shape.equationParagraphs?.[0]?.[0];
	if (equation?.kind === 'equation' && equation.omml)
		return wrapEquationShape(base, { omml: equation.omml, fallback: equation.fallback });
	return `${base}${text}</xdr:${tag}>`;
}

function patchTransform(node: XmlElement, shape: ShapeNode, replaceFrame: boolean): void {
	const props = first(node, shape.kind === 'group' ? 'grpSpPr' : 'spPr', NS.xdr);
	if (!props) return;
	const doc = node.ownerDocument;
	const old = first(props, 'xfrm', NS.a);
	if (replaceFrame || !old) {
		const template = parseXml(`<root xmlns:a="${NS.a}">${transform(shape)}</root>`);
		const xfrm = first(template.documentElement, 'xfrm', NS.a);
		if (!xfrm) return;
		const next = doc.importNode(xfrm, true);
		if (old) props.replaceChild(next, old);
		else props.insertBefore(next, props.firstChild);
		return;
	}
	if (shape.rotation) old.setAttribute('rot', String(Math.round(shape.rotation * 60000)));
	else old.removeAttribute('rot');
	for (const attr of ['flipH', 'flipV'] as const) {
		if (shape[attr]) old.setAttribute(attr, '1');
		else old.removeAttribute(attr);
	}
}

/** Patch only the edited fields of an imported drawing anchor. */
export function patchNativeShape(
	xml: string,
	shape: ShapeNode,
	edit: { transform?: true; text?: true; visibility?: true },
): string {
	const doc = parseXml(xml);
	for (const child of elements(doc.documentElement)) {
		const branches = equationShapeBranches(child);
		const nodes = branches
			? [branches.shape, branches.fallback]
			: child.namespaceURI === NS.xdr && ['sp', 'cxnSp', 'grpSp'].includes(child.localName)
				? [child]
				: [];
		for (const node of nodes) {
			if (!node) continue;
			if (edit.transform) patchTransform(node, shape, false);
			if (edit.visibility) {
				const id = node.getElementsByTagNameNS(NS.xdr, 'cNvPr')[0];
				if (shape.hidden) id?.setAttribute('hidden', '1');
				else id?.removeAttribute('hidden');
			}
			if (edit.text && !shape.equationParagraphs) replaceShapeText(node, shape.text?.text ?? '');
		}
	}
	return buildXml(doc);
}

function replaceShapeText(node: XmlElement, text: string): void {
	const doc = node.ownerDocument;
	const old = first(node, 'txBody', NS.xdr);
	const template = parseXml(
		`<root xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}">${textBody(text)}</root>`,
	);
	const body = first(template.documentElement, 'txBody', NS.xdr);
	if (body && old) {
		for (const paragraph of elements(old).filter(
			(child) => child.namespaceURI === NS.a && child.localName === 'p',
		))
			old.removeChild(paragraph);
		for (const paragraph of elements(body).filter(
			(child) => child.namespaceURI === NS.a && child.localName === 'p',
		))
			old.appendChild(doc.importNode(paragraph, true));
	} else if (body) node.appendChild(doc.importNode(body, true));
}

export function patchShapeText(xml: string, text: string): string {
	const doc = parseXml(xml);
	replaceShapeText(doc.documentElement, text);
	return buildXml(doc.documentElement);
}
