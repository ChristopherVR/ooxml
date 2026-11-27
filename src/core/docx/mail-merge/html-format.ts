// The allow-list of the headless merge's HTML values: which elements are blocks, and the run
// formatting and links an element gives its text. Nothing here resolves or fetches anything.
import { parseCssColorHex, parseCssDeclarations, tokenizeHtml } from '../../text/html-tokens';
import type { TextRun } from '../model';

/** At most this many tokens of one value are read; the rest is kept as text. */
export const MAX_HTML_TOKENS = 20_000;

export type HtmlMergeIssue = 'image-refused' | 'link-refused' | 'nested-table-flattened';

export type RunFormat = Partial<
	Pick<TextRun, 'bold' | 'italic' | 'underline' | 'strike' | 'verticalAlign' | 'color' | 'link'>
>;

export const BLOCKS = new Set([
	'p',
	'div',
	'h1',
	'h2',
	'h3',
	'h4',
	'h5',
	'h6',
	'li',
	'ul',
	'ol',
	'table',
	'tr',
	'td',
	'th',
	'blockquote',
	'pre',
	'section',
	'article',
]);
/** Whether `html` holds no block element, so that it replaces a field in place. */
export const isInlineHtml = (html: string): boolean =>
	!tokenizeHtml(html, MAX_HTML_TOKENS).some(
		(token) => token.kind === 'open' && BLOCKS.has(token.name),
	);

export const VOID = new Set([
	'br',
	'img',
	'hr',
	'input',
	'meta',
	'link',
	'wbr',
	'col',
	'area',
	'source',
]);

function linkOf(href: string | undefined): TextRun['link'] | null | undefined {
	if (href === undefined) return undefined;
	const value = href.trim();
	if (/^#[\w.:-]{1,64}$/.test(value)) return { anchor: value.slice(1) };
	if (/^(https?:|mailto:)/i.test(value) && !/[\s<>"]/.test(value) && value.length <= 2048)
		return { href: value };
	return null;
}

function styleFormat(style: string | undefined): RunFormat {
	if (!style) return {};
	const css = parseCssDeclarations(style);
	const format: RunFormat = {};
	const weight = css.get('font-weight');
	if (weight) format.bold = weight === 'bold' || weight === 'bolder' || Number(weight) >= 600;
	const fontStyle = css.get('font-style');
	if (fontStyle) format.italic = fontStyle === 'italic' || fontStyle === 'oblique';
	const decoration = css.get('text-decoration') ?? css.get('text-decoration-line');
	if (decoration) {
		if (decoration.includes('underline')) format.underline = true;
		if (decoration.includes('line-through')) format.strike = true;
	}
	const color = css.get('color');
	const hex = color && parseCssColorHex(color);
	if (hex) format.color = `#${hex}`;
	return format;
}

export function tagFormat(
	name: string,
	attrs: Map<string, string>,
	issues: Set<HtmlMergeIssue>,
): RunFormat {
	const format = styleFormat(attrs.get('style'));
	if (name === 'b' || name === 'strong' || name === 'th') format.bold = true;
	else if (name === 'i' || name === 'em') format.italic = true;
	else if (name === 'u' || name === 'ins') format.underline = true;
	else if (name === 's' || name === 'strike' || name === 'del') format.strike = true;
	else if (name === 'sup') format.verticalAlign = 'superscript';
	else if (name === 'sub') format.verticalAlign = 'subscript';
	else if (name === 'font') {
		const hex = parseCssColorHex(attrs.get('color') ?? '');
		if (hex) format.color = `#${hex}`;
	} else if (name === 'a') {
		const link = linkOf(attrs.get('href'));
		if (link) format.link = link;
		else if (link === null) issues.add('link-refused');
	}
	return format;
}
