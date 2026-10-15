/** Serializes the OMML object trees this package builds back to Office Math markup. */
import { convertLatexToOmml } from './latex-to-omml';
import { escapeXml } from './omml-mathml-helpers';
import type { OmmlNode } from './omml-node';
import { stripXmlOrderSuffix } from './omml-node';

export const OMML_NAMESPACE = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

const escapeAttribute = (value: string) => escapeXml(value).replace(/"/g, '&quot;');

function element(tag: string, value: unknown): string {
	if (value === undefined || value === null) return '';
	if (Array.isArray(value)) return value.map((item) => element(tag, item)).join('');
	if (typeof value !== 'object') {
		const text = String(value);
		const space = tag === 'm:t' && /^\s|\s$/.test(text) ? ' xml:space="preserve"' : '';
		return `<${tag}${space}>${escapeXml(text)}</${tag}>`;
	}
	const node = value as OmmlNode;
	let attributes = '';
	let content = '';
	for (const key of Object.keys(node)) {
		if (key.startsWith('@_')) {
			attributes += ` ${tag === 'm:t' ? '' : 'm:'}${key.slice(2)}="${escapeAttribute(String(node[key]))}"`;
		} else if (key === '#text') {
			content += escapeXml(String(node[key]));
		} else content += element(stripXmlOrderSuffix(key), node[key]);
	}
	if (tag === 'm:t' && /^\s|\s$/.test(content) && !attributes) attributes = ' xml:space="preserve"';
	return content ? `<${tag}${attributes}>${content}</${tag}>` : `<${tag}${attributes}/>`;
}

/** Markup for a tree whose root key is `m:oMath` or `m:oMathPara`, with the math namespace declared. */
export function ommlToXml(tree: OmmlNode): string {
	const markup = Object.keys(tree)
		.map((key) => element(stripXmlOrderSuffix(key), tree[key]))
		.join('');
	return markup.replace(/^<([\w:]+)/, `<$1 xmlns:m="${OMML_NAMESPACE}"`);
}

/**
 * Office Math markup for LaTeX: an `m:oMathPara` for a display equation, an `m:oMath` for an inline
 * one. Returns undefined for empty input and throws the converter's error for LaTeX it cannot read.
 */
export function latexToOmmlXml(latex: string, display: boolean): string | undefined {
	const tree = convertLatexToOmml(latex);
	const para = tree['m:oMathPara'] as OmmlNode | undefined;
	const math = para?.['m:oMath'];
	if (!math) return undefined;
	return ommlToXml(display ? tree : { 'm:oMath': math });
}
