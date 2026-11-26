import { orderChildren } from './element-order';
import type { Paragraph, ParagraphFormatting } from './model';
import { paragraphJustification } from './paragraph-alignment';
import { EAST_ASIAN_TYPOGRAPHY_KEYS } from './paragraph-styles';
import type { RunFormatting } from './run-style-model';
import { orderParagraphProperties } from './tab-stops';
import { writeParagraphProperties } from './write-paragraph-properties';
import { setRunProperties, setToggle } from './write-run';
import { children, first, makeW, WORD_NS, type XmlElement } from './xml';

export function setStyleAttribute(
	element: XmlElement,
	key: string,
	value: string | number | boolean | undefined,
): void {
	if (value === undefined) element.removeAttributeNS(WORD_NS, key);
	else
		element.setAttributeNS(
			WORD_NS,
			`w:${key}`,
			String(typeof value === 'boolean' ? Number(value) : value),
		);
}

export function setStyleValue(
	parent: XmlElement,
	key: string,
	value: string | number | boolean | undefined,
): void {
	const doc = parent.ownerDocument;
	const old = first(parent, key);
	if (value === undefined && old) parent.removeChild(old);
	if (value === undefined) return;
	const node = old ?? makeW(doc, key);
	setStyleAttribute(node, 'val', value);
	if (!old) parent.appendChild(node);
}

export function styleChild(parent: XmlElement, key: string): XmlElement {
	const old = first(parent, key);
	if (old) return old;
	const node = makeW(parent.ownerDocument, key);
	parent.appendChild(node);
	return node;
}

const asParagraph = (formatting: ParagraphFormatting | undefined): Paragraph => ({
	type: 'paragraph',
	id: 'style',
	runs: [],
	...formatting,
});

export function patchStyleParagraph(
	style: XmlElement,
	next: ParagraphFormatting | undefined,
	prior: ParagraphFormatting | undefined,
	numbering?: { numId: number; level: number },
	priorNumbering?: { numId: number; level: number },
): void {
	const formatChanged = JSON.stringify(next ?? {}) !== JSON.stringify(prior ?? {});
	const numberChanged = JSON.stringify(numbering) !== JSON.stringify(priorNumbering);
	if (!formatChanged && !numberChanged) return;
	const doc = style.ownerDocument;
	const pPr = styleChild(style, 'pPr');
	if (numberChanged) {
		if (numbering) {
			const numPr = styleChild(pPr, 'numPr');
			setStyleValue(numPr, 'ilvl', numbering.level);
			setStyleValue(numPr, 'numId', numbering.numId);
			orderChildren(numPr, ['ilvl', 'numId', 'numberingChange', 'ins']);
		} else {
			for (const node of children(pPr, 'numPr')) pPr.removeChild(node);
		}
	}
	if (formatChanged) {
		const after = asParagraph(next);
		const before = asParagraph(prior);
		if (
			after.align !== before.align ||
			after.justification !== before.justification ||
			after.direction !== before.direction
		)
			setStyleValue(pPr, 'jc', paragraphJustification(after));
		for (const key of [...EAST_ASIAN_TYPOGRAPHY_KEYS, 'suppressAutoHyphens'] as const)
			if (after[key] !== before[key]) setToggle(doc, pPr, key, after[key]);
		writeParagraphProperties(doc, pPr, after, before);
		for (const [key, attr] of [
			['spacingBeforeAuto', 'beforeAutospacing'],
			['spacingAfterAuto', 'afterAutospacing'],
		] as const) {
			if (after[key] === before[key]) continue;
			const spacing = styleChild(pPr, 'spacing');
			setStyleAttribute(spacing, attr, after[key]);
			if (!spacing.attributes.length && !spacing.childNodes.length) pPr.removeChild(spacing);
		}
		if (JSON.stringify(after.frame) !== JSON.stringify(before.frame)) {
			const frame = styleChild(pPr, 'framePr');
			for (const [key, attr] of [
				['widthTwips', 'w'],
				['heightTwips', 'h'],
				['heightRule', 'hRule'],
				['xTwips', 'x'],
				['yTwips', 'y'],
				['xAlign', 'xAlign'],
				['yAlign', 'yAlign'],
				['hAnchor', 'hAnchor'],
				['vAnchor', 'vAnchor'],
				['hSpaceTwips', 'hSpace'],
				['vSpaceTwips', 'vSpace'],
				['wrap', 'wrap'],
			] as const) {
				if (after.frame?.[key] !== before.frame?.[key])
					setStyleAttribute(frame, attr, after.frame?.[key]);
			}
			if (!frame.attributes.length && !frame.childNodes.length) pPr.removeChild(frame);
		}
	}
	orderParagraphProperties(pPr);
	if (!pPr.childNodes.length && !pPr.attributes.length) style.removeChild(pPr);
}

export function patchStyleRun(
	style: XmlElement,
	next: RunFormatting | undefined,
	prior: RunFormatting | undefined,
	eastAsiaLanguage?: string,
	priorLanguage?: string,
): void {
	if (
		JSON.stringify(next ?? {}) === JSON.stringify(prior ?? {}) &&
		eastAsiaLanguage === priorLanguage
	)
		return;
	setRunProperties(
		style.ownerDocument,
		style,
		{ text: '', ...next, ...(eastAsiaLanguage !== undefined && { eastAsiaLanguage }) },
		{ text: '', ...prior, ...(priorLanguage !== undefined && { eastAsiaLanguage: priorLanguage }) },
		false,
	);
}
