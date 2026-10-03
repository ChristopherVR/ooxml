import { NS, buildXml, elements, first, parseXml, type XmlDocument } from '../../xml/index.js';
import type { AppProperties, HeadingPair } from './types.js';

type Kind = 'text' | 'int' | 'bool' | 'pairs' | 'titles';
type AppField = keyof AppProperties;

/** Known fields in the order Excel writes them (the schema is `xsd:all`, so order is free). */
const FIELDS: readonly (readonly [AppField, string, Kind])[] = [
	['template', 'Template', 'text'],
	['totalTime', 'TotalTime', 'int'],
	['application', 'Application', 'text'],
	['docSecurity', 'DocSecurity', 'int'],
	['scaleCrop', 'ScaleCrop', 'bool'],
	['headingPairs', 'HeadingPairs', 'pairs'],
	['titlesOfParts', 'TitlesOfParts', 'titles'],
	['manager', 'Manager', 'text'],
	['company', 'Company', 'text'],
	['linksUpToDate', 'LinksUpToDate', 'bool'],
	['sharedDoc', 'SharedDoc', 'bool'],
	['hyperlinkBase', 'HyperlinkBase', 'text'],
	['hyperlinksChanged', 'HyperlinksChanged', 'bool'],
	['appVersion', 'AppVersion', 'text'],
];

const EMPTY_APP =
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
	`<Properties xmlns="${NS.extendedProperties}" xmlns:vt="${NS.vt}"/>`;

function vectorItems(element: Element): Element[] {
	const vector = first(element, 'vector', NS.vt);
	return vector ? elements(vector) : [];
}

function readPairs(element: Element): HeadingPair[] {
	const variants = vectorItems(element).map((variant) => elements(variant)[0]);
	const out: HeadingPair[] = [];
	for (let i = 0; i + 1 < variants.length; i += 2) {
		const name = variants[i]?.textContent ?? '';
		const count = Number.parseInt(variants[i + 1]?.textContent ?? '', 10);
		out.push({ name, count: Number.isFinite(count) ? count : 0 });
	}
	return out;
}

function readField(element: Element, kind: Kind): AppProperties[AppField] {
	const text = element.textContent ?? '';
	if (kind === 'pairs') return readPairs(element);
	if (kind === 'titles') return vectorItems(element).map((item) => item.textContent ?? '');
	if (kind === 'bool') return text.trim() === 'true' || text.trim() === '1';
	if (kind === 'int') {
		const value = Number.parseInt(text, 10);
		return Number.isFinite(value) ? value : undefined;
	}
	return text === '' ? undefined : text;
}

/** Reads the known fields of `docProps/app.xml`. */
export function parseAppProperties(xml: string | undefined): AppProperties {
	const out: Record<string, unknown> = {};
	if (!xml) return out as AppProperties;
	const root = parseXml(xml, { label: 'app properties' }).documentElement;
	for (const [field, local, kind] of FIELDS) {
		const element = first(root, local, NS.extendedProperties);
		if (!element) continue;
		const value = readField(element, kind);
		if (value !== undefined) out[field] = value;
	}
	return out as AppProperties;
}

const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);

function fill(doc: XmlDocument, element: Element, kind: Kind, value: unknown): void {
	while (element.firstChild) element.removeChild(element.firstChild);
	const vt = (local: string, text?: string) => {
		const node = doc.createElementNS(NS.vt, `vt:${local}`);
		if (text !== undefined) node.textContent = text;
		return node;
	};
	if (kind === 'pairs' || kind === 'titles') {
		const items =
			kind === 'pairs'
				? (value as HeadingPair[]).flatMap((pair) => [
						vt('lpstr', pair.name),
						vt('i4', String(Math.trunc(pair.count))),
					])
				: (value as string[]).map((title) => vt('lpstr', title));
		const vector = vt('vector');
		vector.setAttribute('size', String(items.length));
		vector.setAttribute('baseType', kind === 'pairs' ? 'variant' : 'lpstr');
		for (const item of items) {
			if (kind === 'pairs') {
				const variant = vt('variant');
				variant.appendChild(item);
				vector.appendChild(variant);
			} else vector.appendChild(item);
		}
		element.appendChild(vector);
		return;
	}
	element.textContent = kind === 'bool' ? (value ? 'true' : 'false') : String(value);
}

/**
 * Writes `docProps/app.xml`. With `sourceXml` the source part is patched: an element whose
 * value is unchanged is left byte for byte, a changed one is rewritten, an unset one removed,
 * and every element the model does not know (statistics, `DigSig`, `HLinks`...) is kept.
 */
export function writeAppProperties(props: AppProperties, sourceXml?: string): string {
	const doc = parseXml(sourceXml ?? EMPTY_APP, { label: 'app properties' });
	const root = doc.documentElement;
	const prior = parseAppProperties(sourceXml);
	FIELDS.forEach(([field, local, kind], index) => {
		const value = props[field];
		const existing = first(root, local, NS.extendedProperties);
		if (value === undefined || (kind === 'text' && value === '')) {
			// An empty element (Excel writes `<Company></Company>`) already means unset.
			if (existing && (existing.textContent ?? '') !== '') root.removeChild(existing);
			return;
		}
		if (existing && same(prior[field], value)) return;
		let element = existing;
		if (!element) {
			element = doc.createElementNS(NS.extendedProperties, local);
			const next = FIELDS.slice(index + 1)
				.map(([, name]) => first(root, name, NS.extendedProperties))
				.find((node) => node !== undefined);
			root.insertBefore(element, next ?? null);
		}
		fill(doc, element, kind, value);
	});
	const out = buildXml(doc);
	return out.startsWith('<?xml')
		? out
		: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${out}`;
}
