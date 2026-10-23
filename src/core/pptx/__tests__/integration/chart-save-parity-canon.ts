import { NS, isElement, parseXml, type XmlElement } from '../../../xml/index';

/**
 * Canonical text of an XML part for the chart save harness (`chart-save-parity.test.ts`): one line
 * per element (its path of known-prefix names and its attributes sorted, namespace declarations
 * left out) and one per non-blank text node, so two writers that differ only in prefix choice,
 * declaration placement, attribute order, quoting or empty-element form compare equal.
 */
const PREFIX_BY_URI = new Map<string, string>(
	Object.entries(NS).map(([prefix, uri]) => [uri as string, prefix]),
);

const nameOf = (uri: string | null, local: string): string => {
	if (!uri) return local;
	return `${PREFIX_BY_URI.get(uri) ?? `{${uri}}`}:${local}`;
};

function visit(element: XmlElement, path: string, out: string[]): void {
	const name = nameOf(element.namespaceURI, element.localName);
	const here = `${path}/${name}`;
	const attributes = Array.from(element.attributes)
		.filter((a) => a.prefix !== 'xmlns' && a.name !== 'xmlns')
		.map((a) => `${nameOf(a.namespaceURI, a.localName)}="${a.value}"`)
		.sort();
	out.push(`${here}${attributes.length > 0 ? ` ${attributes.join(' ')}` : ''}`);
	for (const node of Array.from(element.childNodes)) {
		if (isElement(node)) visit(node, here, out);
		else if ((node.nodeType === 3 || node.nodeType === 4) && node.nodeValue?.trim())
			out.push(`${here}#text ${JSON.stringify(node.nodeValue)}`);
	}
}

/** The canonical lines of `xml` (see the module comment). */
export function canonicalXmlLines(xml: string): string[] {
	const out: string[] = [];
	visit(parseXml(xml, { label: 'Chart' }).documentElement, '', out);
	return out;
}

/** A minimal line diff (longest common subsequence): `-` lines only in `a`, `+` only in `b`. */
export function lineDiff(a: readonly string[], b: readonly string[]): string[] {
	const n = a.length;
	const m = b.length;
	const table: Uint32Array[] = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1));
	for (let i = n - 1; i >= 0; i--)
		for (let j = m - 1; j >= 0; j--)
			table[i]![j] =
				a[i] === b[j] ? table[i + 1]![j + 1]! + 1 : Math.max(table[i + 1]![j]!, table[i]![j + 1]!);
	const out: string[] = [];
	let i = 0;
	let j = 0;
	while (i < n && j < m) {
		if (a[i] === b[j]) {
			i++;
			j++;
		} else if (table[i + 1]![j]! >= table[i]![j + 1]!) out.push(`- ${a[i++]}`);
		else out.push(`+ ${b[j++]}`);
	}
	while (i < n) out.push(`- ${a[i++]}`);
	while (j < m) out.push(`+ ${b[j++]}`);
	return out;
}
