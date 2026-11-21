import {
	WORD_ART_FONT,
	WORD_ART_OUTLINE_EMU,
	WORD_ART_SIZE_PT,
	wordArtRgb,
	applyWordArtWarp,
	readWordArtWarp,
	wordArtShadowElement,
	type WordArtText,
} from '../drawingml/word-art';
import { NS, buildXml, elements, first, parseXml, type XmlElement } from '../xml/index';
import type { DrawingAnchor } from './model';
import type { ShapeNode } from './model-shapes';
import { escapeAttr, escapeText } from './write/xml-out';

export function wordArtShape(value: WordArtText, anchor: DrawingAnchor): ShapeNode {
	const color = { kind: 'srgb' as const, value: value.style.fill, transforms: [] };
	return {
		kind: 'shape',
		name: 'WordArt',
		textBox: true,
		wordArt: value.style,
		frame: { x: 0, y: 0, width: anchor.ext?.cx ?? 0, height: anchor.ext?.cy ?? 0 },
		geometry: 'rect',
		fill: { kind: 'none' },
		text: {
			text: value.text,
			anchor: 'ctr',
			paragraphs: [
				{
					align: 'ctr',
					runs: [
						{
							text: value.text,
							sizePt: WORD_ART_SIZE_PT,
							bold: true,
							typeface: WORD_ART_FONT,
							color,
						},
					],
				},
			],
		},
	};
}

/** The simple native WordArt shape written here. Other shapes keep their source markup. */
export function readWordArtShape(body: XmlElement | undefined): ShapeNode['wordArt'] {
	const a = (node: XmlElement | undefined, local: string) => first(node, local, NS.a);
	const bodyProperties = a(body, 'bodyPr');
	if (bodyProperties?.getAttribute('fromWordArt') !== '1') return undefined;
	if (!body) return undefined;
	const paragraphs = elements(body).filter(
		(item) => item.namespaceURI === NS.a && item.localName === 'p',
	);
	const paragraph = paragraphs.length === 1 ? paragraphs[0] : undefined;
	if (!paragraph) return undefined;
	const runs = elements(paragraph).filter(
		(item) => item.namespaceURI === NS.a && item.localName === 'r',
	);
	const run = runs.length === 1 ? runs[0] : undefined;
	const props = a(run, 'rPr');
	const rgb = (parent: XmlElement | undefined) =>
		wordArtRgb(a(a(parent, 'solidFill'), 'srgbClr')?.getAttribute('val'));
	const fill = rgb(props);
	const line = a(props, 'ln');
	const outline = rgb(line);
	const shadow = wordArtRgb(
		a(a(a(props, 'effectLst'), 'outerShdw'), 'srgbClr')?.getAttribute('val'),
	);
	const warp = readWordArtWarp(bodyProperties);
	return fill && (!line || outline)
		? {
				fill,
				...(outline ? { outline } : {}),
				...(shadow ? { shadow } : {}),
				...(warp ? { warp } : {}),
			}
		: undefined;
}

export function wordArtShapeXml(node: ShapeNode, id: number): string {
	const fill = `<a:solidFill><a:srgbClr val="${wordArtRgb(node.wordArt?.fill) ?? '26374A'}"/></a:solidFill>`;
	const outline = wordArtRgb(node.wordArt?.outline);
	const stroke = outline
		? `<a:ln w="${WORD_ART_OUTLINE_EMU}"><a:solidFill><a:srgbClr val="${outline}"/></a:solidFill></a:ln>`
		: '';
	const { width, height } = node.frame;
	const doc = parseXml(`<a:bodyPr xmlns:a="${NS.a}" fromWordArt="1" anchor="ctr"/>`);
	applyWordArtWarp(doc.documentElement!, node.wordArt?.warp);
	const shadow = wordArtRgb(node.wordArt?.shadow);
	const effect = shadow
		? `<a:effectLst>${buildXml(wordArtShadowElement(doc, shadow, NS.a, 'a'))}</a:effectLst>`
		: '';
	return (
		`<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${escapeAttr(node.name ?? 'WordArt')}"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>` +
		`<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${Math.round(width)}" cy="${Math.round(height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></xdr:spPr>` +
		`<xdr:txBody>${buildXml(doc.documentElement!)}<a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="${WORD_ART_SIZE_PT * 100}" b="1">${stroke}${fill}${effect}<a:latin typeface="${WORD_ART_FONT}"/></a:rPr><a:t>${escapeText(node.text?.text ?? '')}</a:t></a:r></a:p></xdr:txBody></xdr:sp>`
	);
}

/** Edits the styled run while retaining the drawing's identity, anchor and extensions. */
export function patchWordArtShape(xml: string, node: ShapeNode): string {
	const doc = parseXml(xml);
	const shape = first(doc.documentElement, 'sp', NS.xdr);
	const body = first(shape, 'txBody', NS.xdr);
	const run = first(first(body, 'p', NS.a), 'r', NS.a);
	const text = first(run, 't', NS.a);
	const props = first(run, 'rPr', NS.a);
	const color = first(first(props, 'solidFill', NS.a), 'srgbClr', NS.a);
	const line = first(props, 'ln', NS.a);
	if (!text || !props || !color || !node.wordArt) return xml;
	const bodyProperties = first(body, 'bodyPr', NS.a);
	if (bodyProperties) applyWordArtWarp(bodyProperties, node.wordArt.warp);
	text.textContent = node.text?.text ?? '';
	color.setAttribute('val', node.wordArt.fill);
	const outline = node.wordArt.outline;
	if (outline) {
		const next = line ?? doc.createElementNS(NS.a, 'a:ln');
		if (!line) props.insertBefore(next, props.firstChild);
		next.setAttribute('w', String(WORD_ART_OUTLINE_EMU));
		let stroke = first(first(next, 'solidFill', NS.a), 'srgbClr', NS.a);
		if (!stroke) {
			const fill = doc.createElementNS(NS.a, 'a:solidFill');
			stroke = doc.createElementNS(NS.a, 'a:srgbClr');
			fill.appendChild(stroke);
			next.appendChild(fill);
		}
		stroke.setAttribute('val', outline);
	} else if (line) {
		props.removeChild(line);
	}
	const effects = first(props, 'effectLst', NS.a);
	const oldShadow = first(effects, 'outerShdw', NS.a);
	const nextShadow = wordArtRgb(node.wordArt.shadow);
	const oldColor = wordArtRgb(first(oldShadow, 'srgbClr', NS.a)?.getAttribute('val'));
	if (nextShadow !== oldColor) {
		if (oldShadow) effects?.removeChild(oldShadow);
		if (nextShadow) {
			const list = effects ?? doc.createElementNS(NS.a, 'a:effectLst');
			if (!effects) props.insertBefore(list, first(props, 'latin', NS.a) ?? null);
			list.appendChild(wordArtShadowElement(doc, nextShadow, NS.a, 'a'));
		}
	}
	return buildXml(doc.documentElement ?? doc);
}
