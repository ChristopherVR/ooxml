/** OMML compatibility tree: attributes use `@_`, text uses `#text`. */
export type OmmlNode = Record<string, unknown>;

/** Retain the established encoded key spelling when returning legacy object trees. */
export function orderedXmlKey(baseName: string, order: number): string {
	return `${baseName}#pptx-order-${order}`;
}

/** Accept both the legacy spelling and the neutral DOM adapter's ordered keys. */
export function stripXmlOrderSuffix(tagName: string): string {
	return tagName.split(/#(?:pptx|math)-order-/u, 1)[0] ?? tagName;
}
