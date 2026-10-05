import { buildXml } from '../xml/index.js';
import { fail, type VisioPackageLimits } from './package-common.js';
import { attribute, children } from './sheet.js';
import { inspectXml, inspectNamespaces } from './xml-validation.js';

export function replacePlainText(
	root: Element,
	shapeId: string,
	text: string,
	check: () => void,
): boolean {
	const shapeChildren = (parent: Element): Element[] => {
		const containers = children(parent, 'Shapes');
		if (containers.length > 1)
			fail('INVALID_SHAPE_ID', 'Duplicate Shapes containers are ambiguous.');
		return children(containers[0], 'Shape');
	};
	const shapes = new Map<string, { node: Element; inherited: boolean; deleted: boolean }>();
	const pending = shapeChildren(root).map((node) => ({
		node,
		inherited: false,
		deleted: false,
	}));
	while (pending.length) {
		check();
		const item = pending.pop()!;
		const id = attribute(item.node, 'ID');
		if (!id || shapes.has(id))
			fail('INVALID_SHAPE_ID', 'Local shape IDs must be present and unique.');
		const inherited =
			item.inherited || item.node.hasAttribute('Master') || item.node.hasAttribute('MasterShape');
		const deleted = item.deleted || ['1', 'true'].includes(attribute(item.node, 'Del') ?? '');
		shapes.set(id, { node: item.node, inherited, deleted });
		for (const node of shapeChildren(item.node)) pending.push({ node, inherited, deleted });
	}
	const target = shapes.get(shapeId);
	if (!target) fail('EDIT_TARGET_NOT_FOUND', 'Local shape does not exist.');
	if (target.inherited || target.deleted)
		fail('UNSUPPORTED_TEXT_EDIT', 'Master-linked or deleted shapes cannot be edited.');
	const texts = children(target.node, 'Text');
	if (texts.length !== 1)
		fail('UNSUPPORTED_TEXT_EDIT', 'Editing requires one existing local Text element.');
	const node = texts[0]!;
	if (Array.from(node.childNodes).some((item) => item.nodeType !== 3 && item.nodeType !== 4))
		fail('UNSUPPORTED_TEXT_EDIT', 'Rich text, fields and unknown text markup cannot be edited.');
	if (children(target.node, 'Section').some((section) => attribute(section, 'N') === 'Field'))
		fail('UNSUPPORTED_TEXT_EDIT', 'Shapes with text fields cannot be edited.');
	if (node.textContent === text) return false;
	while (node.firstChild) node.removeChild(node.firstChild);
	node.appendChild(node.ownerDocument!.createTextNode(text));
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
