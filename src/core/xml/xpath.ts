import * as xpath from 'xpath';
import type { XmlNode } from './xml';

export interface XmlXPathOptions {
	namespaces?: Readonly<Record<string, string>>;
	/** Match unprefixed element names inside a default namespace, as Excel FILTERXML does. */
	matchDefaultNamespace?: boolean;
}

// xpath exposes its compiled-expression API but omits it from the bundled declarations.
interface CompiledXPath {
	select(options: {
		node: XmlNode;
		namespaces?: Readonly<Record<string, string>>;
		allowAnyNamespaceForNoPrefix: boolean;
	}): XmlNode[];
}
const compiler = xpath as typeof xpath & { parse(expression: string): CompiledXPath };

/** XPath 1.0 node selection; invalid expressions and scalar results throw. */
export function selectXmlNodes(
	expression: string,
	node: XmlNode,
	options: XmlXPathOptions = {},
): XmlNode[] {
	return compiler.parse(expression).select({
		node,
		...(options.namespaces ? { namespaces: options.namespaces } : {}),
		allowAnyNamespaceForNoPrefix: options.matchDefaultNamespace ?? false,
	});
}
