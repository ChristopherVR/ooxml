import { getW, isElement, named, WORD_NS, type XmlElement } from './xml';

/** A calculated pagination cache marker, not an authored break. Changed runs invalidate it. */
export function isCachedPageBreak(element: XmlElement): boolean {
	return (
		named(element, 'lastRenderedPageBreak') &&
		!element.childNodes.length &&
		Array.from(element.attributes).every(
			(attribute) => attribute.namespaceURI === 'http://www.w3.org/2000/xmlns/',
		)
	);
}

/** Only ordinary text-wrapping breaks are represented by a model newline. */
export function hasSpecialBreak(element: XmlElement): boolean {
	for (const attribute of Array.from(element.attributes)) {
		if (attribute.namespaceURI === 'http://www.w3.org/2000/xmlns/') continue;
		if (attribute.namespaceURI !== WORD_NS) return true;
		if (attribute.localName === 'type') {
			if (attribute.value !== 'textWrapping') return true;
		} else if (attribute.localName === 'clear') {
			if (attribute.value !== 'none') return true;
		} else return true;
	}
	return false;
}

/** `w:br/@w:type`; absent type defaults to `textWrapping` per the OOXML schema. */
export function classifyBreak(element: XmlElement): 'page' | 'column' | 'textWrapping' {
	const type = getW(element, 'type');
	return type === 'page' || type === 'column' ? type : 'textWrapping';
}

/**
 * A page or column break with no other attributes. These are modeled explicitly (`TextRun.break`)
 * and are safe to relocate on edit; other special breaks (e.g. `w:clear`) remain unsupported.
 */
export function isModeledBreak(element: XmlElement): boolean {
	const kind = classifyBreak(element);
	if (kind === 'textWrapping') return false;
	for (const attribute of Array.from(element.attributes)) {
		if (attribute.namespaceURI === 'http://www.w3.org/2000/xmlns/') continue;
		if (attribute.namespaceURI !== WORD_NS || attribute.localName !== 'type') return false;
	}
	return true;
}

/**
 * Splits a run that holds a page or column break next to other content into one run per piece
 * (each keeps the run's `w:rPr`), so the break is a break run of its own as the model expects.
 * Runs with a single content child, field characters or instructions are returned unchanged.
 */
export function splitBreakRun(run: XmlElement): XmlElement[] {
	const children = Array.from(run.childNodes).filter(isElement);
	const content = children.filter(
		(child) => child.localName !== 'rPr' && !isCachedPageBreak(child),
	);
	if (content.length < 2) return [run];
	const isBreak = (child: XmlElement) => named(child, 'br') && isModeledBreak(child);
	if (!content.some(isBreak)) return [run];
	if (content.some((child) => named(child, 'fldChar') || named(child, 'instrText'))) return [run];
	const props = children.find((child) => child.localName === 'rPr');
	const pieces: XmlElement[] = [];
	let segment: XmlElement[] = [];
	const flush = () => {
		if (!segment.length) return;
		const piece = run.cloneNode(false) as XmlElement;
		if (props) piece.appendChild(props.cloneNode(true));
		for (const child of segment) piece.appendChild(child.cloneNode(true));
		pieces.push(piece);
		segment = [];
	};
	for (const child of content) {
		if (isBreak(child)) flush();
		segment.push(child);
		if (isBreak(child)) flush();
	}
	flush();
	return pieces;
}
