import { buildXml } from '../xml/index';
import { fail, type VisioPackageLimits } from './package-common';
import { attribute, children } from './sheet';
import { inspectXml, inspectNamespaces } from './xml-validation';
import { localTextTarget, textTarget } from './edit-text-target';
import { encodeVisioPlainText, decodeVisioPlainText } from './plain-text';

/** Resolves the master shape a stencil instance (or a sub-shape of one) inherits from. */
export type MasterTemplate = (masterId: string, masterShapeId?: string) => Promise<Element>;

const plain = (node: Element) =>
	Array.from(node.childNodes).every((item) => item.nodeType === 3 || item.nodeType === 4);
const hasFields = (shape: Element) =>
	children(shape, 'Section').some((section) => attribute(section, 'N') === 'Field');
/** A zero, error-free, formula-free cached lock; anything else is treated as protected. */
function assertUnlocked(shape: Element, lock: string): void {
	for (const cell of children(shape, 'Cell')) {
		if (attribute(cell, 'N') !== lock) continue;
		if (cell.hasAttribute('E') || cell.hasAttribute('F') || Number(attribute(cell, 'V')) !== 0)
			fail('EDIT_PROTECTED_CELL', 'The shape is protected against text editing.');
	}
}

/**
 * Text of a stencil (master) instance. An instance inherits its text until it has its own `Text`
 * element; typing in one writes that element and nothing else, which is what Visio saves. Master
 * formulas that size the shape from its text (TEXTHEIGHT of TheText) are recalculated by Visio
 * when it opens the file; the saved size is kept until then.
 */
function replaceInstanceText(shape: Element, template: Element, text: string): boolean {
	const texts = children(shape, 'Text');
	if (texts.length > 1) fail('UNSUPPORTED_TEXT_EDIT', 'The shape has more than one Text element.');
	const local = texts[0];
	const templateTexts = children(template, 'Text');
	if (templateTexts.length > 1)
		fail('UNSUPPORTED_TEXT_EDIT', 'The master shape has more than one Text element.');
	const source = local ?? templateTexts[0];
	if (source && !plain(source))
		fail('UNSUPPORTED_TEXT_EDIT', 'Rich text, fields and unknown text markup cannot be edited.');
	if (hasFields(shape) || hasFields(template))
		fail(
			'UNSUPPORTED_TEXT_EDIT',
			'Replacing all text would remove its text fields; edit the text around the fields instead.',
		);
	// A local lock wins; without one the master's applies.
	const lock = children(shape, 'Cell').some((cell) => attribute(cell, 'N') === 'LockTextEdit')
		? shape
		: template;
	assertUnlocked(lock, 'LockTextEdit');
	if (decodeVisioPlainText(source?.textContent ?? '') === text) return false;
	const document = shape.ownerDocument!;
	let node = local;
	if (!node) {
		node = document.createElementNS(shape.namespaceURI, 'Text');
		// Schema order: cells and sections, then Text, then data, foreign data and sub-shapes.
		const after = Array.from(shape.childNodes).find(
			(child) =>
				child.nodeType === 1 &&
				['Data1', 'Data2', 'Data3', 'ForeignData', 'Shapes'].includes((child as Element).localName),
		);
		shape.insertBefore(node, after ?? null);
	}
	while (node.firstChild) node.removeChild(node.firstChild);
	node.appendChild(document.createTextNode(encodeVisioPlainText(text)));
	return true;
}

export async function replacePlainText(
	root: Element,
	document: Element,
	shapeId: string,
	text: string,
	check: () => void,
	assertDependencies: (shape: Element, text: Element) => Promise<void>,
	masterTemplate?: MasterTemplate,
): Promise<boolean> {
	const found = textTarget(root, shapeId, check);
	if (found.masterId !== undefined) {
		if (!masterTemplate)
			fail('UNSUPPORTED_TEXT_EDIT', 'Stencil shapes cannot be edited without their masters.');
		const template = await masterTemplate(found.masterId, found.masterShapeId);
		check();
		return replaceInstanceText(found.node, template, text);
	}
	const target = localTextTarget(root, document, shapeId, check);
	const texts = children(target, 'Text');
	if (texts.length !== 1)
		fail('UNSUPPORTED_TEXT_EDIT', 'Editing requires one existing local Text element.');
	const node = texts[0]!;
	if (!plain(node))
		fail('UNSUPPORTED_TEXT_EDIT', 'Rich text, fields and unknown text markup cannot be edited.');
	if (hasFields(target))
		fail(
			'UNSUPPORTED_TEXT_EDIT',
			'Replacing all text would remove its text fields; edit the text around the fields instead.',
		);
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
