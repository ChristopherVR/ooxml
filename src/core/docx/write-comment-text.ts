import { diffText } from '../text/text-diff';
import { first, makeW, textContent, WORD_NS, type XmlElement } from './xml';

/** Patch surviving text nodes in place so their run properties, links and surrounding markup stay. */
function patchParagraph(paragraph: XmlElement, text: string): void {
	let nodes = Array.from(paragraph.getElementsByTagNameNS(WORD_NS, 't')) as XmlElement[];
	if (!nodes.length) {
		const run = makeW(paragraph.ownerDocument!, 'r');
		const node = makeW(paragraph.ownerDocument!, 't');
		run.appendChild(node);
		paragraph.appendChild(run);
		nodes = [node];
	}
	const values = nodes.map(textContent);
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
	nodes.forEach((node, index) => {
		node.textContent = values[index]!;
		node.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
	});
}

export function patchCommentText(element: XmlElement, text: string): void {
	const doc = element.ownerDocument!;
	const paragraphs = Array.from(element.getElementsByTagNameNS(WORD_NS, 'p')) as XmlElement[];
	const lines = text.split('\n');
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
