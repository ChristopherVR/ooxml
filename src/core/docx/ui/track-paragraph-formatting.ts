import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';
import { AttrStep, ReplaceAroundStep, type Step } from 'prosemirror-transform';
import { PARAGRAPH_FORMAT_KEYS } from '../restore-paragraph-format';
import { parsePropertiesSnapshot, propertiesSignature } from '../revision-properties';
import { createParagraph } from '../write-paragraph';
import { buildXml, children, first, makeW, parseXml, WORD_NS, type XmlElement } from '../xml';
import { paragraphFromAttrs } from './paragraph-attributes';
import { paragraphFormattingRevision } from './review-paragraph-formatting';

const FORMAT_ATTRS = new Set<string>([...PARAGRAPH_FORMAT_KEYS, 'numId', 'ilvl']);

/** Reject mixed text, mark, structural and document-setting changes before recording history. */
function sameContent(before: ProseMirrorNode, after: ProseMirrorNode): boolean {
	if (before === after) return true;
	if (
		before.type !== after.type ||
		before.text !== after.text ||
		before.childCount !== after.childCount
	)
		return false;
	if (
		JSON.stringify(before.marks.map((mark) => mark.toJSON())) !==
		JSON.stringify(after.marks.map((mark) => mark.toJSON()))
	)
		return false;
	for (const key of new Set([...Object.keys(before.attrs), ...Object.keys(after.attrs)])) {
		if (before.type.name === 'paragraph' && FORMAT_ATTRS.has(key)) continue;
		if (JSON.stringify(before.attrs[key]) !== JSON.stringify(after.attrs[key])) return false;
	}
	for (let index = 0; index < before.childCount; index++)
		if (!sameContent(before.child(index), after.child(index))) return false;
	return true;
}

/** Paragraph-mark and section properties have independent histories. */
function paragraphProperties(element: XmlElement): XmlElement {
	for (const name of ['rPr', 'sectPr', 'pPrChange'])
		for (const child of children(element, name)) element.removeChild(child);
	return element;
}
function properties(node: ProseMirrorNode): XmlElement {
	const paragraph = paragraphFromAttrs(node.attrs, String(node.attrs.id), []);
	delete paragraph.formatRevision;
	delete paragraph.markRevision;
	const doc = parseXml(`<w:p xmlns:w="${WORD_NS}"/>`);
	return paragraphProperties(
		first(createParagraph(doc, paragraph, undefined), 'pPr') ?? makeW(doc, 'pPr'),
	);
}

/** Record pure paragraph formatting through the shared writer and immutable prior snapshot. */
export function trackParagraphFormatting(
	steps: readonly Step[],
	oldState: EditorState,
	newState: EditorState,
	author: string,
	date: string,
	nextId: () => string,
): Transaction | null {
	if (
		!steps.length ||
		!steps.every((step) => step instanceof AttrStep || step instanceof ReplaceAroundStep) ||
		!sameContent(oldState.doc, newState.doc)
	)
		return null;
	const tr = newState.tr;
	let id: string | undefined;
	newState.doc.descendants((node, pos) => {
		if (node.type.name !== 'paragraph' || !node.type.spec.attrs?.formatRevision) return;
		const previous = oldState.doc.nodeAt(pos)!;
		if (previous === node) return;
		const before = properties(previous);
		const after = properties(node);
		if (propertiesSignature(before) === propertiesSignature(after)) return;
		const pending = paragraphFormattingRevision(previous);
		const priorXml = pending?.previousParagraphPropertiesXml ?? buildXml(before);
		if (!priorXml) return;
		const restored =
			propertiesSignature(after) ===
			propertiesSignature(paragraphProperties(parsePropertiesSnapshot(priorXml, 'pPr')));
		tr.setNodeMarkup(pos, undefined, {
			...node.attrs,
			formatRevision: restored
				? null
				: (pending ?? {
						kind: 'paragraphChange',
						author,
						date,
						dateUtc: date,
						id: (id ??= nextId()),
						previousParagraphPropertiesXml: priorXml,
					}),
		});
	});
	return tr.docChanged ? tr : null;
}
