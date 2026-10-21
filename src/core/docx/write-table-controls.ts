// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Keeps the content controls around table rows and cells in step with the model: an edited control
// is patched in place, a control the model dropped is unwrapped (its row or cell stays), and a
// control the model added wraps its row or cell in a new `w:sdt`.
import type { ContentControl } from './content-control-model';
import { isSdt } from './content-control';
import {
	createContentControl,
	patchContentControl,
	unwrapContentControl,
} from './write-content-control';
import { isElement, named, type XmlDocument, type XmlElement } from './xml';

/** The `w:sdt` whose content directly holds `element`, if any. */
function enclosingSdt(element: XmlElement): XmlElement | undefined {
	const content = element.parentNode as XmlElement | null;
	const sdt = content?.parentNode as XmlElement | null | undefined;
	return named(content, 'sdtContent') && isSdt(sdt) ? sdt : undefined;
}

/** Replaces `sdt` by its content, when `element` is the only row or cell it holds. */
function unwrap(sdt: XmlElement, element: XmlElement): void {
	const content = element.parentNode as XmlElement;
	const others = Array.from(content.childNodes)
		.filter(isElement)
		.some((child) => child !== element && (named(child, 'tc') || named(child, 'tr')));
	if (!others) unwrapContentControl(sdt);
}

/**
 * Writes the control of the row or cell `element` (already in place in its table): `next` is the
 * model's control, `base` the one it was loaded with. An unchanged control is left as it was.
 */
export function placeTableControl(
	doc: XmlDocument,
	element: XmlElement,
	next: ContentControl | undefined,
	base: ContentControl | undefined,
): void {
	if (JSON.stringify(next) === JSON.stringify(base)) return;
	const sdt = enclosingSdt(element);
	if (next && sdt) patchContentControl(doc, sdt, next);
	else if (sdt) unwrap(sdt, element);
	else if (next && element.parentNode) {
		const created = createContentControl(doc, next);
		element.parentNode.replaceChild(created.sdt, element);
		created.content.appendChild(element);
	}
}
