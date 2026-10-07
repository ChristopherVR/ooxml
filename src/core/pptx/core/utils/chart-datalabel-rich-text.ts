/**
 * Builds the `c:tx` of a data label with literal text, keeping the run
 * formatting PowerPoint authored.
 *
 * @module utils/chart-datalabel-rich-text
 */

import type { XmlObject } from '../types';
import type { GetLocalName } from './chart-title-xml-ops';
import { collectAllText, findKey } from './chart-title-xml-ops';

/** The first element of a child that may repeat. */
function firstNode(value: unknown): XmlObject | undefined {
	const item = Array.isArray(value) ? value[0] : value;
	return item && typeof item === 'object' ? (item as XmlObject) : undefined;
}

/** `parent`'s first child with the local name `local`. */
function child(
	parent: XmlObject,
	local: string,
	getLocalName: GetLocalName,
): XmlObject | undefined {
	const key = findKey(parent, local, getLocalName);
	return key ? firstNode(parent[key]) : undefined;
}

/**
 * The run properties edited label text should keep: the first run's, else the
 * first field's, else the paragraph's end-of-paragraph properties.
 */
function labelRunProperties(
	paragraph: XmlObject,
	getLocalName: GetLocalName,
): XmlObject | undefined {
	for (const local of ['r', 'fld']) {
		const runNode = child(paragraph, local, getLocalName);
		const rPr = runNode ? child(runNode, 'rPr', getLocalName) : undefined;
		if (rPr) {
			return rPr;
		}
	}
	return child(paragraph, 'endParaRPr', getLocalName);
}

/**
 * The `c:tx` for a label with literal text.
 *
 * An authored `c:rich` whose text is unchanged is kept as it is. When the text
 * was edited, it is written as one run that keeps the first run's `a:rPr`, and
 * the first paragraph keeps its `a:pPr` and `a:endParaRPr`, so the label keeps
 * its font, size and colour. A label with no `c:rich` gets a plain one.
 */
export function buildDataLabelTx(
	existingTx: XmlObject | undefined,
	text: string,
	getLocalName: GetLocalName,
): XmlObject {
	const richKey = existingTx ? findKey(existingTx, 'rich', getLocalName) : undefined;
	const rich = richKey ? firstNode(existingTx?.[richKey]) : undefined;
	if (!existingTx || !richKey || !rich) {
		return {
			'c:rich': { 'a:bodyPr': {}, 'a:lstStyle': {}, 'a:p': { 'a:r': { 'a:t': text } } },
		};
	}
	const authored: string[] = [];
	collectAllText(rich, getLocalName, authored);
	if (authored.join('') === text) {
		return existingTx;
	}
	const pKey = findKey(rich, 'p', getLocalName) ?? 'a:p';
	const paragraph = firstNode(rich[pKey]) ?? {};
	const pPr = child(paragraph, 'pPr', getLocalName);
	const rPr = labelRunProperties(paragraph, getLocalName);
	const endParaRPr = child(paragraph, 'endParaRPr', getLocalName);
	const run: XmlObject = rPr ? { 'a:rPr': rPr, 'a:t': text } : { 'a:t': text };
	const newParagraph: XmlObject = {};
	if (pPr) {
		newParagraph['a:pPr'] = pPr;
	}
	newParagraph['a:r'] = run;
	if (endParaRPr) {
		newParagraph['a:endParaRPr'] = endParaRPr;
	}
	const newRich: XmlObject = { ...rich };
	newRich[pKey] = newParagraph;
	const tx: XmlObject = { ...existingTx };
	tx[richKey] = newRich;
	return tx;
}
