import type { XmlNode } from './xml';

export interface XmlXPathOptions {
	namespaces?: Readonly<Record<string, string>>;
	/** Match unprefixed element names inside a default namespace, as Excel FILTERXML does. */
	matchDefaultNamespace?: boolean;
}

interface CompiledXPath {
	select(options: {
		node: XmlNode;
		namespaces?: Readonly<Record<string, string>>;
		allowAnyNamespaceForNoPrefix: boolean;
	}): XmlNode[];
}

export type XPathSelector = (
	expression: string,
	node: XmlNode,
	options?: XmlXPathOptions,
) => XmlNode[];

export function xpathSelector(api: typeof import('xpath')): XPathSelector {
	// xpath exposes the compiled-expression API but omits it from its declarations.
	const compiler = api as typeof api & { parse(expression: string): CompiledXPath };
	return (expression, node, options = {}) =>
		compiler.parse(expression).select({
			node,
			...(options.namespaces ? { namespaces: options.namespaces } : {}),
			allowAnyNamespaceForNoPrefix: options.matchDefaultNamespace ?? false,
		});
}
