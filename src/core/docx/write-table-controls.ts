// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Keeps the content controls around table rows and cells in step with the model: an edited control
// is patched in place, a control the model dropped from every row or cell it holds is unwrapped
// (they stay), and a control the model added wraps its row or cell in a new `w:sdt`.
import type { ContentControl } from './content-control-model';
import { isSdt } from './content-control';
import {
	createContentControl,
	patchContentControl,
	unwrapContentControl,
} from './write-content-control';
import { named, type XmlDocument, type XmlElement } from './xml';

/** The `w:sdt` whose content directly holds `element`, if any. */
function enclosingSdt(element: XmlElement): XmlElement | undefined {
	const content = element.parentNode as XmlElement | null;
	const sdt = content?.parentNode as XmlElement | null | undefined;
	return named(content, 'sdtContent') && isSdt(sdt) ? sdt : undefined;
}

/**
 * Writes the control of the row or cell `element` (already in place in its table): `next` is the
 * model's control, `base` the one it was loaded with. An unchanged control is left as it was; a
 * dropped one is unwrapped by `unwrapDroppedControls`, which sees every row or cell it holds.
 */
export function placeTableControl(
	doc: XmlDocument,
	element: XmlElement,
	next: ContentControl | undefined,
	base: ContentControl | undefined,
): void {
	if (!next || JSON.stringify(next) === JSON.stringify(base)) return;
	const sdt = enclosingSdt(element);
	if (sdt) patchContentControl(doc, sdt, next);
	else if (element.parentNode) {
		const created = createContentControl(doc, next);
		element.parentNode.replaceChild(created.sdt, element);
		created.content.appendChild(element);
	}
}

/**
 * Unwraps each `w:sdt` around `elements` (the rows of a table or the cells of a row, in place)
 * that the model takes off every one of them it holds: `next` is each element's control. The
 * rows or cells stay; a control the model keeps on any of them stays around all of them.
 */
export function unwrapDroppedControls(
	elements: readonly XmlElement[],
	next: readonly (ContentControl | undefined)[],
): void {
	const kept = new Map<XmlElement, boolean>();
	elements.forEach((element, index) => {
		const sdt = enclosingSdt(element);
		if (sdt) kept.set(sdt, Boolean(kept.get(sdt)) || next[index] !== undefined);
	});
	for (const [sdt, keep] of kept) if (!keep) unwrapContentControl(sdt);
}
