// Raw XML fragments of a chart part (`c:spPr` children, text bodies, extension lists): serialised
// once by the parser and re-emitted by the writer. The serializer repeats the declaration of every
// prefix a fragment uses; the writer drops those the part root already declares, so a fragment
// written back under the same root reproduces the source bytes.
import { NS, buildXml, elements, type XmlElement } from '../xml/index';

/** Prefix -> namespace bindings a fragment is written under. */
export type NamespaceBindings = ReadonlyMap<string, string>;

/** The bindings every chart part root declares (`c`, `a`, `r`). */
export const CHART_ROOT_BINDINGS: NamespaceBindings = new Map([
	['c', NS.c],
	['a', NS.a],
	['r', NS.r],
]);

const DECLARATION = / xmlns:([A-Za-z_][\w.-]*)="([^"]*)"/g;

/**
 * Removes the declarations of `xml` that repeat a binding in `bindings`. A prefix the fragment also
 * binds to another namespace is left alone, so no rebinding inside the fragment changes meaning.
 */
export function stripDeclarations(xml: string, bindings: NamespaceBindings): string {
	const conflicting = new Set<string>();
	for (const [, prefix, uri] of xml.matchAll(DECLARATION))
		if (prefix !== undefined && bindings.has(prefix) && bindings.get(prefix) !== uri)
			conflicting.add(prefix);
	return xml.replace(DECLARATION, (match, prefix: string, uri: string) =>
		bindings.get(prefix) === uri && !conflicting.has(prefix) ? '' : match,
	);
}

/** The child elements of `element`, serialised, with the chart root's declarations dropped. */
export const innerXml = (element: XmlElement): string =>
	stripDeclarations(elements(element).map(buildXml).join(''), CHART_ROOT_BINDINGS);
