import {
	WORD_ART_FONT,
	WORD_ART_OUTLINE_EMU,
	WORD_ART_SIZE_PT,
	wordArtRgb,
	readWordArtFill,
	readWordArtGlow,
	readWordArtReflection,
	wordArtFillElement,
	wordArtGlowElement,
	wordArtPaintEquals,
	wordArtReflectionElement,
	readWordArtWarp,
	wordArtShadowElement,
	type WordArtStyle,
} from '../drawingml/word-art';
import { makeNS, makeW, type XmlDocument, type XmlElement } from './xml';

const WORD_2010 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const WORD_MAIN = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export function wordArtRunProperties(doc: XmlDocument, style: WordArtStyle): XmlElement {
	const props = makeW(doc, 'rPr');
	const wordProp = (name: string, attributes: Record<string, string> = {}) => {
		const node = makeW(doc, name);
		for (const [attribute, value] of Object.entries(attributes))
			node.setAttributeNS(node.namespaceURI, `w:${attribute}`, value);
		props.appendChild(node);
	};
	const extension = (name: string) => makeNS(doc, WORD_2010, `w14:${name}`);
	const solid = (color: string) => {
		const fill = extension('solidFill');
		const rgb = extension('srgbClr');
		rgb.setAttributeNS(WORD_2010, 'w14:val', color);
		fill.appendChild(rgb);
		return fill;
	};
	wordProp('rFonts', { ascii: WORD_ART_FONT, hAnsi: WORD_ART_FONT });
	wordProp('b');
	const fill = wordArtRgb(style.fill) ?? '26374A';
	wordProp('color', { val: fill });
	wordProp('sz', { val: String(WORD_ART_SIZE_PT * 2) });
	const stroke = wordArtRgb(style.outline);
	const shadow = wordArtRgb(style.shadow);
	if (shadow) props.appendChild(wordArtShadowElement(doc, shadow, WORD_2010, 'w14'));
	if (style.glow) props.appendChild(wordArtGlowElement(doc, style.glow, WORD_2010, 'w14'));
	if (style.reflection)
		props.appendChild(wordArtReflectionElement(doc, style.reflection, WORD_2010, 'w14'));
	if (stroke) {
		const outline = extension('textOutline');
		outline.setAttributeNS(WORD_2010, 'w14:w', String(WORD_ART_OUTLINE_EMU));
		outline.appendChild(solid(stroke));
		props.appendChild(outline);
	}
	const textFill = extension('textFill');
	textFill.appendChild(wordArtFillElement(doc, style, WORD_2010, 'w14'));
	props.appendChild(textFill);
	return props;
}

export function readWordArtStyle(box: XmlElement | undefined): WordArtStyle | undefined {
	if (!box) return undefined;
	const bodies = box.getElementsByTagNameNS(
		'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
		'bodyPr',
	);
	if (bodies[0]?.getAttribute('fromWordArt') !== '1') return undefined;
	const fills = Array.from(box.getElementsByTagNameNS(WORD_2010, 'textFill'));
	const colors = Array.from(box.getElementsByTagNameNS(WORD_MAIN, 'color'));
	const outlines = Array.from(box.getElementsByTagNameNS(WORD_2010, 'textOutline'));
	const shadows = Array.from(box.getElementsByTagNameNS(WORD_2010, 'shadow'));
	const glows = Array.from(box.getElementsByTagNameNS(WORD_2010, 'glow'));
	const reflections = Array.from(box.getElementsByTagNameNS(WORD_2010, 'reflection'));
	const rgb = (parent: XmlElement | undefined) =>
		wordArtRgb(
			parent?.getElementsByTagNameNS(WORD_2010, 'srgbClr')[0]?.getAttributeNS(WORD_2010, 'val'),
		);
	// Desktop Word drops a duplicate textFill when w:color already stores the same solid color.
	const plainRgb = (node: XmlElement | undefined) =>
		wordArtRgb(node?.getAttributeNS(WORD_MAIN, 'val'));
	const plain = plainRgb(colors[0]);
	const paint = fills.length
		? readWordArtFill(fills[0], WORD_2010)
		: plain
			? { fill: plain }
			: undefined;
	const outline = rgb(outlines[0]);
	const shadow = rgb(shadows[0]);
	const glow = readWordArtGlow(glows[0], WORD_2010);
	const reflection = readWordArtReflection(reflections[0], WORD_2010);
	if (
		!paint ||
		fills.some((node) => {
			const item = readWordArtFill(node, WORD_2010);
			return !item || !wordArtPaintEquals(item, paint);
		}) ||
		(fills.length === 0 && colors.some((node) => plainRgb(node) !== paint.fill)) ||
		outlines.some((node) => rgb(node) !== outline) ||
		shadows.some((node) => rgb(node) !== shadow) ||
		glows.some(
			(node) => JSON.stringify(readWordArtGlow(node, WORD_2010)) !== JSON.stringify(glow),
		) ||
		reflections.some(
			(node) =>
				JSON.stringify(readWordArtReflection(node, WORD_2010)) !== JSON.stringify(reflection),
		)
	) {
		return undefined;
	}
	const warp = readWordArtWarp(bodies[0]);
	return {
		...paint,
		...(outline ? { outline } : {}),
		...(shadow ? { shadow } : {}),
		...(glow ? { glow } : {}),
		...(reflection ? { reflection } : {}),
		...(warp ? { warp } : {}),
	};
}
