// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Data-bound content controls (`w:dataBinding`): Word shows the value of the custom XML node the
// binding's XPath selects, not the text in the document, so an edit of the control's text must be
// written to that node as well. Supported XPath: an absolute path of child steps (`/a/b`,
// `/ns0:a/ns0:b[2]`), optionally ending in an attribute (`/@id`) or `text()`. Anything else is
// refused with a diagnostic, since writing a guessed node would show a wrong value in Word.
import type JSZip from 'jszip';
import type { Block, DocumentModel, Paragraph } from './model';
import { buildXml, elements, first, getW, parseXml, WORD_NS } from './xml';
import type { XmlElement } from './xml';

interface Binding {
	xpath: string;
	prefixes: Map<string, string>;
	storeItemId: string;
}

interface Step {
	prefix: string;
	name: string;
	index: number;
}

/** `xmlns:p='uri'` declarations of `w:prefixMappings`. */
function parsePrefixes(text: string | undefined): Map<string, string> {
	const map = new Map<string, string>();
	for (const match of (text ?? '').matchAll(/xmlns:([\w.-]+)\s*=\s*(?:'([^']*)'|"([^"]*)")/g))
		map.set(match[1]!, match[2] ?? match[3] ?? '');
	return map;
}

/** Bindings of the source document's controls, by control id (`w:dataBinding`, `w15:dataBinding`). */
function sourceBindings(sourceXml: string): Map<string, Binding> {
	const result = new Map<string, Binding>();
	const document = parseXml(sourceXml);
	for (const sdt of Array.from(document.getElementsByTagNameNS(WORD_NS, 'sdt'))) {
		const sdtPr = first(sdt, 'sdtPr');
		const element = sdtPr && elements(sdtPr).find((node) => node.localName === 'dataBinding');
		const id = getW(first(sdtPr, 'id'), 'val');
		const xpath = element?.getAttributeNS(WORD_NS, 'xpath');
		const storeItemId = element?.getAttributeNS(WORD_NS, 'storeItemID');
		if (id && element && xpath && storeItemId)
			result.set(id, {
				xpath,
				storeItemId,
				prefixes: parsePrefixes(element.getAttributeNS(WORD_NS, 'prefixMappings') ?? undefined),
			});
	}
	return result;
}

function paragraphsOf(blocks: readonly Block[]): Paragraph[] {
	return blocks.flatMap((block) =>
		block.type === 'paragraph'
			? [block]
			: block.rows.flatMap((row) => row.flatMap((c) => c.paragraphs)),
	);
}

/** The text of each data-bound control of the body, by id; block controls join paragraphs by newline. */
function boundTexts(model: DocumentModel): Map<string, string> {
	const parts = new Map<string, string[]>();
	const add = (id: string, text: string) => parts.set(id, [...(parts.get(id) ?? []), text]);
	for (const paragraph of paragraphsOf(model.blocks)) {
		const own = paragraph.contentControl;
		if (own?.dataBound) {
			add(own.id, paragraph.runs.map((run) => run.text).join(''));
			continue;
		}
		const inline = new Map<string, string>();
		for (const run of paragraph.runs) {
			const control = run.contentControl;
			if (!control?.dataBound) continue;
			if (!inline.has(control.id)) inline.set(control.id, '');
			if (!control.showingPlaceholder) inline.set(control.id, inline.get(control.id)! + run.text);
		}
		for (const [id, text] of inline) add(id, text);
	}
	return new Map([...parts].map(([id, texts]) => [id, texts.join('\n')]));
}

function parseXPath(binding: Binding): { steps: Step[]; attribute?: Step } | undefined {
	const parts = binding.xpath.split('/');
	if (parts.shift() !== '' || !parts.length) return undefined;
	if (parts.at(-1) === 'text()') parts.pop();
	const parse = (raw: string, attribute: boolean): Step | undefined => {
		const match = /^(?:([A-Za-z_][\w.-]*):)?([A-Za-z_][\w.-]*)(?:\[(\d+)\])?$/.exec(raw);
		if (!match || (attribute && match[3])) return undefined;
		const [, prefix = '', name = '', index] = match;
		if (prefix && !binding.prefixes.has(prefix)) return undefined;
		return { prefix, name, index: index ? Number(index) : 1 };
	};
	const last = parts.at(-1);
	const attribute = last?.startsWith('@') ? parse(last.slice(1), true) : undefined;
	if (last?.startsWith('@')) {
		parts.pop();
		if (!attribute) return undefined;
	}
	const steps = parts.map((part) => parse(part, false));
	if (!steps.length || steps.some((step) => !step)) return undefined;
	return { steps: steps as Step[], ...(attribute ? { attribute } : {}) };
}

const matches = (element: XmlElement, step: Step, binding: Binding): boolean =>
	element.localName === step.name &&
	(element.namespaceURI ?? '') === (step.prefix ? (binding.prefixes.get(step.prefix) ?? '') : '');

/** Sets the text of the node `binding` selects in `root`; false when the path selects nothing. */
function writeValue(root: XmlElement, binding: Binding, value: string): boolean | 'unsupported' {
	const path = parseXPath(binding);
	if (!path) return 'unsupported';
	let node: XmlElement | undefined = root;
	for (const [depth, step] of path.steps.entries()) {
		const candidates: XmlElement[] =
			depth === 0
				? [root].filter((e) => matches(e, step, binding))
				: elements(node!).filter((e) => matches(e, step, binding));
		node = candidates[step.index - 1];
		if (!node) return false;
	}
	if (path.attribute) {
		const { prefix, name } = path.attribute;
		const ns = prefix ? binding.prefixes.get(prefix) : undefined;
		if (ns) node!.setAttributeNS(ns, `${prefix}:${name}`, value);
		else node!.setAttribute(name, value);
		return true;
	}
	for (const child of Array.from(node!.childNodes)) node!.removeChild(child);
	if (value) node!.appendChild(node!.ownerDocument.createTextNode(value));
	return true;
}

/** The custom XML item part whose `ds:itemID` is `storeItemId`, if the package has one. */
async function findItemPart(zip: JSZip, storeItemId: string): Promise<string | undefined> {
	const wanted = storeItemId.toLowerCase();
	for (const name of Object.keys(zip.files)) {
		const props = /^customXml\/itemProps(\d+)\.xml$/.exec(name);
		if (!props) continue;
		const root = parseXml(await zip.file(name)!.async('string')).documentElement;
		const id = Array.from(root.attributes).find((attribute) => attribute.localName === 'itemID');
		if ((id?.value ?? '').toLowerCase() === wanted) return `customXml/item${props[1]}.xml`;
	}
	return undefined;
}

/**
 * Writes the changed text of data-bound controls to their custom XML parts. A control whose store
 * part the package does not hold is left alone (there is no value to keep in step); an XPath
 * outside the supported forms, or one that selects no node, throws.
 */
export async function applyDataBindings(
	zip: JSZip,
	model: DocumentModel,
	base: DocumentModel,
	sourceXml: string,
): Promise<void> {
	const now = boundTexts(model);
	if (!now.size) return;
	const before = boundTexts(base);
	const bindings = sourceBindings(sourceXml);
	const parsed = new Map<string, ReturnType<typeof parseXml>>();
	for (const [id, text] of now) {
		const binding = bindings.get(id);
		if (!binding || before.get(id) === text) continue;
		const part = await findItemPart(zip, binding.storeItemId);
		const file = part ? zip.file(part) : undefined;
		if (!part || !file) continue;
		const document = parsed.get(part) ?? parseXml(await file.async('string'));
		parsed.set(part, document);
		const written = writeValue(document.documentElement, binding, text);
		if (written !== true)
			throw new Error(
				written === 'unsupported'
					? `Cannot update the data-bound content control ${id}: the XPath ${binding.xpath} is not supported, so Word would keep showing the old bound value. The original DOCX package remains unchanged.`
					: `Cannot update the data-bound content control ${id}: the XPath ${binding.xpath} selects no node in ${part}. The original DOCX package remains unchanged.`,
			);
	}
	for (const [part, document] of parsed) zip.file(part, buildXml(document));
}
