import { diffText } from '../text/text-diff';
import { first, makeW, named, WORD_NS, type XmlElement } from './xml';
import { paragraphText, paragraphTextNodes } from './paragraph-text';
import { runContentText } from './special-characters';
import { createRun } from './write-run';

/** Keep an unchanged special character's original element, including break/tab attributes. */
function replaceTextNode(node: XmlElement, value: string): void {
	const old = runContentText(node)!;
	if (old === value) return;
	if (named(node, 't') && !/[\t\n\u00ad\u2011]/.test(value)) {
		node.textContent = value;
		node.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
		return;
	}
	const parent = node.parentNode!;
	const insert = (text: string, before: ChildNode | null) => {
		if (!text) return;
		const run = createRun(node.ownerDocument!, { text });
		for (const child of Array.from(run.childNodes)) parent.insertBefore(child, before);
	};
	const kept = named(node, 't') || !old ? -1 : value.indexOf(old);
	if (kept >= 0) {
		insert(value.slice(0, kept), node);
		insert(value.slice(kept + old.length), node.nextSibling);
	} else {
		insert(value, node);
		parent.removeChild(node);
	}
}

/** Patch surviving text nodes in place so their run properties, links and surrounding markup stay. */
function patchParagraph(paragraph: XmlElement, text: string): void {
	let nodes = paragraphTextNodes(paragraph);
	if (!nodes.length) {
		const run = makeW(paragraph.ownerDocument!, 'r');
		const node = makeW(paragraph.ownerDocument!, 't');
		run.appendChild(node);
		paragraph.appendChild(run);
		nodes = [node];
	}
	const values = nodes.map((node) => runContentText(node)!);
	const lengths = values.map((value) => value.length);
	const old = values.join('');
	// Apply changes backwards: offsets into the original styled nodes remain valid.
	for (const change of diffText(old, text).reverse()) {
		let offset = 0;
		let inserted = false;
		for (let index = 0; index < nodes.length; index++) {
			const value = values[index]!;
			const end = offset + lengths[index]!;
			if (end >= change.fromStart && offset <= change.fromEnd) {
				const from = Math.max(0, change.fromStart - offset);
				const to = Math.min(lengths[index]!, change.fromEnd - offset);
				const replacement = inserted ? '' : text.slice(change.toStart, change.toEnd);
				values[index] = value.slice(0, from) + replacement + value.slice(Math.max(from, to));
				inserted = true;
			}
			offset = end;
		}
	}
	nodes.forEach((node, index) => replaceTextNode(node, values[index]!));
}

/** Existing soft breaks remain inline; newly inserted newlines make comment paragraphs. */
function commentParagraphs(paragraphs: XmlElement[], text: string): string[] {
	const before = paragraphs.map(paragraphText);
	const changes = diffText(before.join('\n'), text);
	const softBreaks = new Set<number>();
	let offset = 0;
	for (const paragraph of before) {
		for (const match of paragraph.matchAll(/\n/g)) {
			const position = offset + match.index;
			let shift = 0;
			let removed = false;
			for (const change of changes) {
				if (position < change.fromStart) break;
				if (position < change.fromEnd) {
					removed = true;
					break;
				}
				shift = change.toEnd - change.fromEnd;
			}
			if (!removed) softBreaks.add(position + shift);
		}
		offset += paragraph.length + 1;
	}
	const lines: string[] = [];
	let start = 0;
	for (const match of text.matchAll(/\n/g)) {
		if (softBreaks.has(match.index)) continue;
		lines.push(text.slice(start, match.index));
		start = match.index + 1;
	}
	lines.push(text.slice(start));
	return lines;
}

export function patchCommentText(element: XmlElement, text: string): void {
	const doc = element.ownerDocument!;
	const paragraphs = Array.from(element.getElementsByTagNameNS(WORD_NS, 'p')) as XmlElement[];
	const lines = commentParagraphs(paragraphs, text);
	for (let index = 0; index < Math.max(lines.length, paragraphs.length); index++) {
		let paragraph = paragraphs[index];
		if (!paragraph) {
			paragraph = makeW(doc, 'p');
			const properties = first(paragraphs.at(-1), 'pPr');
			if (properties) paragraph.appendChild(properties.cloneNode(true));
			element.appendChild(paragraph);
		}
		if (index >= lines.length && paragraph.parentNode === element) element.removeChild(paragraph);
		else patchParagraph(paragraph, lines[index] ?? '');
	}
}
