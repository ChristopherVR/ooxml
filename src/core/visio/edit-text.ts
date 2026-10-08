import { buildXml } from '../xml/index';
import { fail, type VisioPackageLimits } from './package-common';
import { attribute, children } from './sheet';
import { inspectXml, inspectNamespaces } from './xml-validation';
import { localTextTarget } from './edit-text-target';
import { encodeVisioPlainText, decodeVisioPlainText } from './plain-text';

export async function replacePlainText(
	root: Element,
	document: Element,
	shapeId: string,
	text: string,
	check: () => void,
	assertDependencies: (shape: Element, text: Element) => Promise<void>,
): Promise<boolean> {
	const target = localTextTarget(root, document, shapeId, check);
	const texts = children(target, 'Text');
	if (texts.length !== 1)
		fail('UNSUPPORTED_TEXT_EDIT', 'Editing requires one existing local Text element.');
	const node = texts[0]!;
	if (Array.from(node.childNodes).some((item) => item.nodeType !== 3 && item.nodeType !== 4))
		fail('UNSUPPORTED_TEXT_EDIT', 'Rich text, fields and unknown text markup cannot be edited.');
	if (children(target, 'Section').some((section) => attribute(section, 'N') === 'Field'))
		fail('UNSUPPORTED_TEXT_EDIT', 'Shapes with text fields cannot be edited.');
	if (decodeVisioPlainText(node.textContent ?? '') === text) return false;
	await assertDependencies(target, node);
	check();
	while (node.firstChild) node.removeChild(node.firstChild);
	node.appendChild(node.ownerDocument!.createTextNode(encodeVisioPlainText(text)));
	return true;
}

export function serializeEditedXml(
	root: Element,
	limits: VisioPackageLimits,
	check: () => void,
): { bytes: Uint8Array; nodes: number } {
	check();
	// XMLSerializer writes normalized whitespace literally; reject instead of silently
	// changing character-reference CR text or tab/newline attribute values on reopen.
	const pending: Node[] = [root.ownerDocument!];
	while (pending.length) {
		check();
		const node = pending.pop()!;
		if ((node.nodeType === 3 || node.nodeType === 4) && node.nodeValue?.includes('\r'))
			fail('UNSUPPORTED_XML_EDIT', 'Carriage-return text cannot be preserved by this writer.');
		if (node.nodeType === 1)
			for (const attr of Array.from((node as Element).attributes))
				if (/[\t\r\n]/.test(attr.value))
					fail('UNSUPPORTED_XML_EDIT', 'Attribute whitespace cannot be preserved by this writer.');
		for (const child of Array.from(node.childNodes)) pending.push(child);
	}
	let xml = buildXml(root.ownerDocument!);
	// Serialization produces UTF-8 bytes, including for a UTF-16 source document.
	xml = xml.replace(/^<\?xml\s[^?]*\?>/, (declaration) =>
		declaration.replace(/encoding\s*=\s*(['"])[^'"]*\1/, 'encoding="UTF-8"'),
	);
	const nodes = inspectXml(xml, limits, check);
	inspectNamespaces(root, check);
	const bytes = new TextEncoder().encode(xml);
	if (bytes.length > limits.maxEntryBytes)
		fail('LIMIT_ENTRY', 'Edited XML part exceeds byte limit.');
	return { bytes, nodes };
}
