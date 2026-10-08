import { NS, elements, first, parseXml, type XmlElement } from '../../xml/index';
import { PartPatch, escapeText } from './patch';
import type { AppProperties, HeadingPair } from './types';

type Kind = 'text' | 'int' | 'bool' | 'pairs' | 'titles';
type AppField = keyof AppProperties;

/**
 * Known fields in one order consistent with the order Excel, Word and PowerPoint each write
 * them; a field the source lacks is inserted before the next known one it has (the schema is
 * `xsd:all`, so any order is valid).
 */
const FIELDS: readonly (readonly [AppField, string, Kind])[] = [
	['template', 'Template', 'text'],
	['totalTime', 'TotalTime', 'int'],
	['pages', 'Pages', 'int'],
	['words', 'Words', 'int'],
	['characters', 'Characters', 'int'],
	['application', 'Application', 'text'],
	['docSecurity', 'DocSecurity', 'int'],
	['presentationFormat', 'PresentationFormat', 'text'],
	['lines', 'Lines', 'int'],
	['paragraphs', 'Paragraphs', 'int'],
	['slides', 'Slides', 'int'],
	['notes', 'Notes', 'int'],
	['hiddenSlides', 'HiddenSlides', 'int'],
	['mmClips', 'MMClips', 'int'],
	['scaleCrop', 'ScaleCrop', 'bool'],
	['headingPairs', 'HeadingPairs', 'pairs'],
	['titlesOfParts', 'TitlesOfParts', 'titles'],
	['manager', 'Manager', 'text'],
	['company', 'Company', 'text'],
	['linksUpToDate', 'LinksUpToDate', 'bool'],
	['charactersWithSpaces', 'CharactersWithSpaces', 'int'],
	['sharedDoc', 'SharedDoc', 'bool'],
	['hyperlinkBase', 'HyperlinkBase', 'text'],
	['hyperlinksChanged', 'HyperlinksChanged', 'bool'],
	['appVersion', 'AppVersion', 'text'],
];

/** ISO/IEC 29500 Strict names of the extended-properties and variant-type namespaces. */
export const STRICT_APP_NAMESPACES = {
	extendedProperties: 'http://purl.oclc.org/ooxml/officeDocument/extendedProperties',
	vt: 'http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes',
} as const;

const EMPTY_APP =
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
	`<Properties xmlns="${NS.extendedProperties}" xmlns:vt="${NS.vt}"/>`;

/** The part's namespaces: Transitional, or Strict when the root is in the Strict namespace. */
function namespacesOf(root: XmlElement): { app: string; vt: string } | undefined {
	if (root.localName !== 'Properties') return undefined;
	if (root.namespaceURI === NS.extendedProperties) return { app: NS.extendedProperties, vt: NS.vt };
	if (root.namespaceURI === STRICT_APP_NAMESPACES.extendedProperties)
		return { app: STRICT_APP_NAMESPACES.extendedProperties, vt: STRICT_APP_NAMESPACES.vt };
	return undefined;
}

function vectorItems(element: Element, vt: string): Element[] {
	const vector = first(element, 'vector', vt);
	return vector ? elements(vector) : [];
}

function readPairs(element: Element, vt: string): HeadingPair[] {
	const variants = vectorItems(element, vt).map((variant) => elements(variant)[0]);
	const out: HeadingPair[] = [];
	for (let i = 0; i + 1 < variants.length; i += 2) {
		const name = variants[i]?.textContent ?? '';
		const count = Number.parseInt(variants[i + 1]?.textContent ?? '', 10);
		out.push({ name, count: Number.isFinite(count) ? count : 0 });
	}
	return out;
}

function readField(element: Element, kind: Kind, vt: string): AppProperties[AppField] {
	const text = element.textContent ?? '';
	if (kind === 'pairs') return readPairs(element, vt);
	if (kind === 'titles') return vectorItems(element, vt).map((item) => item.textContent ?? '');
	if (kind === 'bool') return text.trim() === 'true' || text.trim() === '1';
	if (kind === 'int') {
		const value = Number.parseInt(text, 10);
		return Number.isFinite(value) ? value : undefined;
	}
	return text === '' ? undefined : text;
}

function readKnown(root: XmlElement, ns: { app: string; vt: string }): AppProperties {
	const out: Record<string, unknown> = {};
	for (const [field, local, kind] of FIELDS) {
		const element = first(root, local, ns.app);
		if (!element) continue;
		const value = readField(element, kind, ns.vt);
		if (value !== undefined) out[field] = value;
	}
	return out as AppProperties;
}

/** Reads the known fields of `docProps/app.xml` (Transitional or Strict). */
export function parseAppProperties(xml: string | undefined): AppProperties {
	if (!xml) return {};
	const root = parseXml(xml, { label: 'app properties' }).documentElement;
	const ns = namespacesOf(root);
	return ns ? readKnown(root, ns) : {};
}

const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);

function content(kind: Kind, value: unknown, vt: string): string {
	if (kind === 'bool') return value ? 'true' : 'false';
	if (kind === 'int') return String(Math.trunc(value as number));
	if (kind === 'text') return escapeText(String(value));
	const item = (local: string, text: string) =>
		`<${vt}:${local}>${escapeText(text)}</${vt}:${local}>`;
	const items =
		kind === 'pairs'
			? (value as HeadingPair[]).flatMap((pair) => [
					`<${vt}:variant>${item('lpstr', pair.name)}</${vt}:variant>`,
					`<${vt}:variant>${item('i4', String(Math.trunc(pair.count)))}</${vt}:variant>`,
				])
			: (value as string[]).map((title) => item('lpstr', title));
	const baseType = kind === 'pairs' ? 'variant' : 'lpstr';
	return `<${vt}:vector size="${items.length}" baseType="${baseType}">${items.join('')}</${vt}:vector>`;
}

/**
 * Writes `docProps/app.xml`. With `sourceXml` the source part is patched: an element whose
 * value is unchanged is left byte for byte, a changed one has its content rewritten, an unset
 * one is removed, and everything the model does not know (`DigSig`, `HLinks`...) is kept
 * exactly (see `PartPatch`). A Strict source is written in its own namespaces.
 */
export function writeAppProperties(props: AppProperties, sourceXml?: string): string {
	const xml = sourceXml ?? EMPTY_APP;
	const root = parseXml(xml, { label: 'app properties' }).documentElement;
	const ns = namespacesOf(root);
	if (!ns) throw new Error('app properties root is not an extended-properties element');
	const patch = new PartPatch(xml, root);
	const prior = readKnown(root, ns);
	const ownPrefix = root.prefix ? `${root.prefix}:` : '';
	FIELDS.forEach(([field, local, kind], index) => {
		const value = props[field];
		const existing = first(root, local, ns.app);
		if (value === undefined || (kind === 'text' && value === '')) {
			// An empty element (Excel writes `<Company></Company>`) already means unset.
			if (existing && (existing.textContent ?? '') !== '') patch.remove(existing);
			return;
		}
		if (existing && same(prior[field], value)) return;
		const vt = kind === 'pairs' || kind === 'titles' ? patch.prefix(ns.vt, 'vt') : '';
		const body = content(kind, value, vt);
		if (existing) {
			patch.setContent(existing, body);
			return;
		}
		const next = FIELDS.slice(index + 1)
			.map(([, name]) => first(root, name, ns.app))
			.find((node) => node !== undefined);
		patch.insert(`<${ownPrefix}${local}>${body}</${ownPrefix}${local}>`, next);
	});
	return patch.toString();
}
