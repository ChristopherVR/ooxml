// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Reads a source paragraph's inline content back into the slots the model's runs came from, and
// the position markers between them, so the writer can rebuild the paragraph without losing XML.
import { alternateContentDrawing } from './write-drawing';
import { isSdt, sdtContent } from './content-control';
import { isEquationElement } from './equation';
import { isModeledFormFieldData } from './form-fields';
import { annotationEdge } from './inline-annotations';
import { isAnnotationWrapper, isPositionMarker, type InlineMarker } from './inline-markers';
import { runPositionalTab, runSymbol } from './special-characters';
import { hasSpecialBreak, isCachedPageBreak, isModeledBreak } from './breaks';
import { isCommentReferenceRun } from './parse-revisions';
import { isCommentAnchorElement, isRevisionWrapperElement } from './write-revisions';
import { getW, isElement, named, type XmlElement } from './xml';

/**
 * One modeled run's source element, with the hyperlink or simple field that wrapped it (if any).
 * A simple field without a cached result has a slot with no element.
 */
export interface InlineSlot {
	element?: XmlElement;
	container?: XmlElement;
	/** The tracked change (`w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`) the run was read in. */
	revision?: XmlElement;
}

/** A run's own content is safe when it is plain text, a modeled break, or a lone picture. */
function runIsSafe(run: XmlElement): boolean {
	const content = Array.from(run.childNodes).filter(
		(child): child is XmlElement =>
			isElement(child) && !named(child, 'rPr') && !isCachedPageBreak(child),
	);
	const only = content.length === 1 ? content[0] : undefined;
	if (only && named(only, 'br') && isModeledBreak(only)) return true;
	// Note references and a note's own number mark are modeled runs (see block-parser.ts).
	if (
		only &&
		['footnoteReference', 'endnoteReference', 'footnoteRef', 'endnoteRef'].some((name) =>
			named(only, name),
		)
	)
		return true;
	// A complex field's marker, with legacy form field data on a begin marker, or its instruction
	// text are modeled runs.
	if (only && named(only, 'fldChar')) {
		const type = getW(only, 'fldCharType') ?? '';
		// Indentation between the marker's tags is not content.
		const data = Array.from(only.childNodes).filter(
			(node) => isElement(node) || Boolean(node.textContent?.trim()),
		);
		if (!data.length) return ['begin', 'separate', 'end'].includes(type);
		const [ffData] = data;
		return (
			type === 'begin' && data.length === 1 && isElement(ffData!) && isModeledFormFieldData(ffData)
		);
	}
	if (content.length && content.every((element) => named(element, 'instrText'))) return true;
	// One kind of symbol and one kind of positional tab per run, as the model holds them.
	if (runSymbol(content) === null || runPositionalTab(content) === null) return false;
	const kinds = new Set<string>();
	for (const element of content) {
		if (element.localName === 'br' && hasSpecialBreak(element)) return false;
		if (element.localName === 'cr' && element.attributes.length > 0) return false;
		// A shape or text box Word wrote for old readers too: its `mc:Choice` drawing is the model.
		if (element.localName === 'AlternateContent') {
			if (!alternateContentDrawing(element)) return false;
			kinds.add('drawing');
			continue;
		}
		const allowed = [
			't',
			'tab',
			'br',
			'cr',
			'noBreakHyphen',
			'softHyphen',
			'sym',
			'ptab',
			'delText',
			'drawing',
			'pict',
		];
		if (!allowed.includes(element.localName)) return false;
		kinds.add(element.localName);
	}
	// A run is modeled as text or as a picture, never both.
	return !((kinds.has('drawing') || kinds.has('pict')) && kinds.size > 1);
}

/**
 * Collects the slots and markers of `parent`'s children: the paragraph itself, or a container
 * (hyperlink, simple field or control content) or wrapper inside it, within the tracked change
 * `revision` if any. Inline content controls are read only at the paragraph's level.
 */
function runsOf(
	parent: XmlElement,
	container: XmlElement | undefined,
	slots: InlineSlot[],
	markers: InlineMarker[],
	paragraph: XmlElement,
	revision?: XmlElement,
): boolean {
	const inContainer = { ...(container && { container }), ...(revision && { revision }) };
	for (const child of Array.from(parent.childNodes)) {
		if (!isElement(child)) {
			if (child.nodeType === 3 && child.textContent?.trim()) return false;
			continue;
		}
		if (parent === paragraph && named(child, 'pPr')) continue;
		if (isEquationElement(child)) {
			slots.push({ element: child, ...inContainer });
		} else if (named(child, 'r')) {
			if (isCommentReferenceRun(child)) continue;
			if (!runIsSafe(child)) return false;
			slots.push({ element: child, ...inContainer });
		} else if (isPositionMarker(child)) {
			markers.push({ element: child, boundary: slots.length, ...inContainer });
		} else if (named(child, 'smartTagPr') || named(child, 'customXmlPr')) {
			// An annotation wrapper's own properties, copied when the wrapper is restored.
			if (!isAnnotationWrapper(parent)) return false;
		} else if (isAnnotationWrapper(child)) {
			const edge = (side: 'start' | 'end') =>
				markers.push({
					element: annotationEdge(child),
					boundary: slots.length,
					annotation: { wrapper: child, edge: side },
					...inContainer,
				});
			edge('start');
			if (!runsOf(child, container, slots, markers, paragraph, revision)) return false;
			edge('end');
		} else if (isRevisionWrapperElement(child) && (!container || isSdt(container))) {
			if (!runsOf(child, container, slots, markers, paragraph, child)) return false;
		} else if (isSdt(child) && parent === paragraph) {
			const content = sdtContent(child);
			const before = slots.length;
			if (content && !runsOf(content, child, slots, markers, paragraph)) return false;
			if (slots.length === before) slots.push({ container: child });
		} else if ((named(child, 'hyperlink') || named(child, 'fldSimple')) && !container) {
			const before = slots.length;
			if (!runsOf(child, child, slots, markers, paragraph, revision)) return false;
			// The parser shows a placeholder for a simple field without a cached result.
			if (named(child, 'fldSimple') && slots.length === before) slots.push({ container: child });
		} else if (!isCommentAnchorElement(child) || container) return false;
	}
	return true;
}

/**
 * Collects the paragraph's modeled runs in the same order the parser produced them, and the
 * position markers between them. Returns `undefined` when the paragraph holds inline content this
 * editor cannot safely regenerate. `w:pPr` is handled by the caller.
 */
export function collectInlineSlots(
	paragraph: XmlElement,
): { slots: InlineSlot[]; markers: InlineMarker[] } | undefined {
	const slots: InlineSlot[] = [];
	const markers: InlineMarker[] = [];
	return runsOf(paragraph, undefined, slots, markers, paragraph) ? { slots, markers } : undefined;
}

/** Top-level paragraph children that `buildInlineContent` output replaces. */
export function replaceableInlineChildren(paragraph: XmlElement): XmlElement[] {
	return Array.from(paragraph.childNodes).filter(
		(child): child is XmlElement =>
			isElement(child) &&
			(isEquationElement(child) ||
				named(child, 'r') ||
				named(child, 'hyperlink') ||
				named(child, 'fldSimple') ||
				isSdt(child) ||
				isRevisionWrapperElement(child) ||
				isAnnotationWrapper(child) ||
				isPositionMarker(child) ||
				isCommentAnchorElement(child)),
	);
}
