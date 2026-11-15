/**
 * Normalises the HTML Word puts on the clipboard into plain HTML the editor's parser reads:
 * `mso-list` paragraphs become list paragraphs (`data-num-id`, `data-ilvl`) with their marker
 * text removed, `class=MsoXxx` and heading tags become paragraph style references, Word-only run
 * properties become `data-run-props`, fields keep their codes (`data-field`), and conditional
 * comments, namespaced tags and `mso-` CSS are dropped. DOM-free: string in, string out.
 */
import { parseDeclarations, serializeDeclarations } from './css';
import {
	decodeEntities,
	escapeText,
	startTag,
	tokenizeHtml,
	tokensText,
	type HtmlToken,
} from './html-tokens';
import { clipboardFragment, isWordClipboardHtml } from './word-fragment';
import { elementEnd, listReference, scanWordLists, type WordClipboardList } from './word-list-scan';
import { wordRunStyle, type PastedRunProperties } from './word-runs';
import { characterStyleName, paragraphStyleName, readWordStyleSheet } from './word-styles';
import { FieldTracker } from './word-fields';

/**
 * Word's Paste Options: `source` keeps the copied styles and direct formatting; `merge` keeps
 * lists and emphasis (bold, italic, underline, strikethrough, super- and subscript) and lets the
 * text take the formatting of the place it is pasted.
 */
export type WordPasteMode = 'source' | 'merge';

export interface WordClipboardOptions {
	mode?: WordPasteMode;
	/** The `numId` to give one pasted list (called once per list), or `undefined` for none. */
	numId?(list: WordClipboardList): number | undefined;
	/** The id of the destination's paragraph style named `name`, or `undefined` when it has none. */
	paragraphStyleId?(name: string): string | undefined;
	/** The id of the destination's character style named `name`. */
	characterStyleId?(name: string): string | undefined;
}

const PARAGRAPHS = /^(p|h[1-6])$/;
const EMPHASIS = new Set(['b', 'strong', 'i', 'em', 'u', 's', 'strike', 'sup', 'sub']);
const RUNS = new Set([...EMPHASIS, 'span', 'font']);
const STRUCTURE = new Set(['table', 'thead', 'tbody', 'tfoot', 'tr', 'td', 'th', 'ul', 'ol', 'li']);
const DROPPED = new Set(['style', 'script', 'title', 'xml', 'head', 'object', 'iframe']);
/** Downlevel-revealed content that only stands in for what the editor reads elsewhere. */
const KEPT_CONDITIONS = new Set(['!vml']);

interface Open {
	/** The source tag the entry closes on. */
	source: string;
	/** The end tags written for it (empty when the element was unwrapped). */
	closing: string;
	props: PastedRunProperties;
}

/** The paragraph attributes the editor reads, for one Word paragraph. */
function paragraphAttributes(
	token: Extract<HtmlToken, { kind: 'open' }>,
	sheet: ReturnType<typeof readWordStyleSheet>,
	options: WordClipboardOptions,
	numIds: Map<string, number | undefined>,
): Record<string, string | undefined> {
	const declarations = parseDeclarations(token.attrs.get('style'));
	const align = declarations.get('text-align') ?? token.attrs.get('align');
	const reference = listReference(token.attrs.get('style'));
	const numId = reference ? numIds.get(reference.key) : undefined;
	const styleName =
		options.mode === 'merge'
			? undefined
			: paragraphStyleName(sheet, token.attrs.get('class'), token.name);
	const style = styleName ? options.paragraphStyleId?.(styleName) : undefined;
	return {
		style:
			align && /^(left|right|center|justify)$/i.test(align)
				? `text-align:${align.toLowerCase()}`
				: undefined,
		dir: token.attrs.get('dir'),
		'data-docx-paragraph-style': style,
		'data-num-id': numId === undefined ? undefined : String(numId),
		'data-ilvl': numId === undefined || !reference ? undefined : String(reference.level),
	};
}

/**
 * The markup written for one run element. Formatting the editor reads from spans only (fonts,
 * colour, highlight, run properties) goes on a span, inside the emphasis tag when there is one.
 */
function runElement(
	token: Extract<HtmlToken, { kind: 'open' }>,
	inherited: PastedRunProperties,
	sheet: ReturnType<typeof readWordStyleSheet>,
	options: WordClipboardOptions,
): { opening: string; closing: string; props: PastedRunProperties } {
	const { tag, attrs, props } = runTag(token, inherited, sheet, options);
	const span = Object.values(attrs).some(Boolean);
	const tags = [...(tag ? [tag] : []), ...(span ? ['span'] : [])];
	return {
		opening: (tag ? `<${tag}>` : '') + (span ? startTag('span', attrs) : ''),
		closing: tags
			.reverse()
			.map((name) => `</${name}>`)
			.join(''),
		props,
	};
}

function runTag(
	token: Extract<HtmlToken, { kind: 'open' }>,
	inherited: PastedRunProperties,
	sheet: ReturnType<typeof readWordStyleSheet>,
	options: WordClipboardOptions,
): { tag: string | null; attrs: Record<string, string | undefined>; props: PastedRunProperties } {
	const name =
		token.name === 'strong'
			? 'b'
			: token.name === 'em'
				? 'i'
				: token.name === 'strike'
					? 's'
					: token.name;
	const run = wordRunStyle(parseDeclarations(token.attrs.get('style')));
	if (options.mode === 'merge') {
		const tag = EMPHASIS.has(name) && !(name === 's' && run.dropStrike) ? name : null;
		return { tag, attrs: {}, props: inherited };
	}
	const own = Object.keys(run.props).length > 0;
	const props = own ? { ...inherited, ...run.props } : inherited;
	const characterName = characterStyleName(sheet, token.attrs.get('class'));
	const characterStyle = characterName ? options.characterStyleId?.(characterName) : undefined;
	const css = serializeDeclarations(run.css);
	const attrs = {
		style: css || undefined,
		lang: token.attrs.get('lang'),
		'data-run-props': own ? JSON.stringify(props) : undefined,
		'data-docx-style': characterStyle,
	};
	const tag = EMPHASIS.has(name) && !(name === 's' && run.dropStrike) ? name : null;
	return { tag, attrs, props };
}

/** Whether a Word footnote or endnote reference anchor (`mso-footnote-id:ftn1`). */
function noteAnchor(token: Extract<HtmlToken, { kind: 'open' }>): boolean {
	const declarations = parseDeclarations(token.attrs.get('style'));
	return declarations.has('mso-footnote-id') || declarations.has('mso-endnote-id');
}

function linkAttributes(
	token: Extract<HtmlToken, { kind: 'open' }>,
): Record<string, string> | null {
	const href = token.attrs.get('href') ?? '';
	// Bookmarks and links into Word's own clip files (`#_ftn1`, `#_Toc...`) lead nowhere here.
	return /^(https?:|mailto:|ftp:|tel:)/i.test(href)
		? { href, title: token.attrs.get('title') ?? '' }
		: null;
}

/**
 * Word's clipboard HTML as the editor's HTML, or `null` when `html` did not come from Word (the
 * caller then pastes it unchanged).
 */
export function normalizeWordClipboardHtml(
	html: string,
	options: WordClipboardOptions = {},
): string | null {
	if (!isWordClipboardHtml(html)) return null;
	const sheet = readWordStyleSheet(html);
	const tokens = tokenizeHtml(clipboardFragment(html));
	const numIds = new Map<string, number | undefined>();
	for (const [key, list] of scanWordLists(tokens, sheet)) numIds.set(key, options.numId?.(list));
	let out = '';
	const stack: Open[] = [];
	const hidden: boolean[] = [];
	const fields = new FieldTracker();
	const props = (): PastedRunProperties => stack.at(-1)?.props ?? {};
	const close = (entry: Open) => {
		out += entry.closing;
	};
	for (let index = 0; index < tokens.length; index++) {
		const token = tokens[index]!;
		if (token.kind === 'marker') {
			if (token.condition === null) hidden.pop();
			else
				hidden.push(
					hidden.includes(true) || !KEPT_CONDITIONS.has(token.condition.replace(/\s+/g, '')),
				);
			continue;
		}
		if (hidden.includes(true)) continue;
		if (token.kind === 'comment') {
			out += fields.comment(token.text, stack.length);
			continue;
		}
		if (token.kind === 'text') {
			if (fields.inCode) fields.addCode(decodeEntities(token.text));
			else out += token.text;
			continue;
		}
		if (token.kind === 'close') {
			const at = stack.map((entry) => entry.source).lastIndexOf(token.name);
			if (at < 0) continue;
			out += fields.closeTo(at);
			while (stack.length > at) close(stack.pop()!);
			continue;
		}
		const { name } = token;
		const declarations = parseDeclarations(token.attrs.get('style'));
		if (DROPPED.has(name) || declarations.get('mso-list')?.toLowerCase() === 'ignore') {
			index = elementEnd(tokens, index) - 1;
			continue;
		}
		if (name === 'a' && noteAnchor(token)) {
			// The reference's number; the note itself follows as text (footnote-list below).
			const end = elementEnd(tokens, index);
			const number = /\d+/.exec(tokensText(tokens.slice(index + 1, end)))?.[0] ?? '*';
			out += `<sup>${escapeText(number)}</sup>`;
			index = end - 1;
			continue;
		}
		if (name === 'br' || name === 'img' || name === 'hr') {
			if (name === 'br') out += '<br>';
			if (name === 'img')
				out += startTag('img', {
					src: token.attrs.get('src'),
					alt: token.attrs.get('alt'),
					width: token.attrs.get('width'),
					height: token.attrs.get('height'),
				});
			continue;
		}
		if (token.selfClosing) continue;
		let entry: Open;
		if (name === 'o:p') {
			// `<o:p>&nbsp;</o:p>` is how Word keeps an empty paragraph from collapsing.
			const end = elementEnd(tokens, index);
			const inner = tokens.slice(index + 1, end - 1);
			if (inner.every((t) => t.kind === 'text' && /^(\s|&nbsp;)*$/.test(t.text))) {
				index = end - 1;
				continue;
			}
			entry = { source: name, closing: '', props: props() };
		} else if (PARAGRAPHS.test(name)) {
			out += startTag('p', paragraphAttributes(token, sheet, options, numIds));
			entry = { source: name, closing: '</p>', props: {} };
		} else if (RUNS.has(name)) {
			const run = runElement(token, props(), sheet, options);
			out += run.opening;
			entry = { source: name, closing: run.closing, props: run.props };
		} else if (name === 'a') {
			const link = linkAttributes(token);
			if (link) out += startTag('a', link);
			entry = { source: name, closing: link ? '</a>' : '', props: props() };
		} else if (STRUCTURE.has(name)) {
			out += startTag(name, {
				colspan: token.attrs.get('colspan'),
				rowspan: token.attrs.get('rowspan'),
			});
			entry = { source: name, closing: `</${name}>`, props: {} };
		} else {
			// Word's wrappers (`div`, `font`, `o:`/`w:`/`v:` tags) carry nothing the editor reads.
			entry = { source: name, closing: '', props: props() };
		}
		stack.push(entry);
	}
	out += fields.closeTo(0);
	while (stack.length) close(stack.pop()!);
	return out.trim();
}
