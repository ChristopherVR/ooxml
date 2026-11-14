// Clipboard HTML (`text/html`, CF_HTML) as DrawingML paragraphs (`a:p`): what a slide, a chart
// title or any other DrawingML text body receives when rich text is pasted into it. The input is
// treated as untrusted: it is lexed without a DOM, nothing in it runs, its size, nesting, token,
// text and paragraph counts are capped, links pass the shared hyperlink policy, and markup the
// model has no place for degrades to its text.
import { hyperlinkPolicy } from '../opc/safe-href';
import { htmlClipboardFragment, tokenizeHtml } from '../text/html-tokens';
import {
	blockAlignment,
	bulletFromMarker,
	cssRunProperties,
	listBullet,
	msoListLevel,
	tagRunProperties,
} from './html-text-style';
import type {
	DrawingParagraph,
	DrawingParagraphProperties,
	DrawingRunProperties,
} from './text-model';

/** The bounds one paste is held to. */
export interface HtmlTextPasteLimits {
	/** Characters of HTML read; the rest is ignored. */
	maxInputLength: number;
	/** Tags and text items lexed; the rest is read as tag-free text. */
	maxTokens: number;
	/** Open elements tracked; deeper elements keep their text but add no formatting. */
	maxDepth: number;
	/** Characters of text produced. */
	maxTextLength: number;
	/** Paragraphs produced. */
	maxParagraphs: number;
}

export const DEFAULT_HTML_TEXT_PASTE_LIMITS: Readonly<HtmlTextPasteLimits> = {
	maxInputLength: 2_000_000,
	maxTokens: 50_000,
	maxDepth: 64,
	maxTextLength: 100_000,
	maxParagraphs: 2_000,
};

/** Converted paragraphs. A run's `hyperlinkClick.relId` names an entry of `hyperlinks`. */
export interface HtmlTextPasteResult {
	paragraphs: DrawingParagraph[];
	/** Placeholder relationship id to the vetted external target; the caller makes the relationship. */
	hyperlinks: Record<string, string>;
	/** A limit cut the input short. */
	truncated: boolean;
}

const BLOCKS = new Set([
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
	'dl',
	'dt',
	'dd',
	'blockquote',
	'pre',
	'address',
	'article',
	'aside',
	'section',
	'header',
	'footer',
	'nav',
	'main',
	'figure',
	'figcaption',
	'center',
	'table',
	'tbody',
	'thead',
	'tfoot',
	'tr',
	'caption',
	'body',
	'html',
]);
const VOID = new Set([
	'br',
	'img',
	'hr',
	'meta',
	'link',
	'input',
	'wbr',
	'col',
	'area',
	'base',
	'source',
	'embed',
	'param',
	'track',
]);

interface Frame {
	name: string;
	run: DrawingRunProperties;
	pre: boolean;
	/** The text of this element is a typed list marker (Word's `mso-list:Ignore` span). */
	marker: boolean;
	align?: string;
	ordered?: boolean;
	attrs?: Map<string, string>;
	msoLevel?: number;
}

const BULLET_INDENT_EMU = 285750;
const LEVEL_STEP_EMU = 457200;

const sameProps = (a?: DrawingRunProperties, b?: DrawingRunProperties): boolean =>
	JSON.stringify(a ?? {}) === JSON.stringify(b ?? {});

/** Converts clipboard HTML into DrawingML paragraphs within `limits`. */
export function htmlToDrawingParagraphs(
	html: string,
	limits: Partial<HtmlTextPasteLimits> = {},
): HtmlTextPasteResult {
	const bounds = { ...DEFAULT_HTML_TEXT_PASTE_LIMITS, ...limits };
	let truncated = html.length > bounds.maxInputLength;
	const source = htmlClipboardFragment(truncated ? html.slice(0, bounds.maxInputLength) : html);
	const tokens = tokenizeHtml(source, bounds.maxTokens);
	if (tokens.length >= bounds.maxTokens) truncated = true;

	const paragraphs: DrawingParagraph[] = [];
	const hyperlinks: Record<string, string> = {};
	const stack: Frame[] = [{ name: '#root', run: {}, pre: false, marker: false }];
	let current: DrawingParagraph | null = null;
	let textLength = 0;
	let marker: string | null = null;
	let conditionalMarker = false;
	let stopped = false;
	const top = (): Frame => stack[stack.length - 1] as Frame;

	const paragraphProperties = (): DrawingParagraphProperties | undefined => {
		const props: DrawingParagraphProperties = {};
		let listDepth = 0;
		let item: Frame | undefined;
		let list: Frame | undefined;
		let msoLevel: number | undefined;
		for (let i = stack.length - 1; i >= 0; i--) {
			const frame = stack[i] as Frame;
			if (props.align === undefined && frame.align) props.align = frame.align;
			if (frame.name === 'li' && !item) item = frame;
			if (frame.name === 'ul' || frame.name === 'ol') {
				listDepth++;
				list ??= frame;
			}
			if (msoLevel === undefined && frame.msoLevel !== undefined) msoLevel = frame.msoLevel;
		}
		if (item && list) {
			const level = Math.min(8, listDepth - 1);
			props.bullet = listBullet(list.ordered ?? false, list.attrs ?? new Map(), level);
			if (level > 0) props.level = level;
			props.marginLeftEmu = BULLET_INDENT_EMU + level * LEVEL_STEP_EMU;
			props.indentEmu = -BULLET_INDENT_EMU;
		} else if (msoLevel !== undefined) {
			if (msoLevel > 0) props.level = msoLevel;
			props.marginLeftEmu = BULLET_INDENT_EMU + msoLevel * LEVEL_STEP_EMU;
			props.indentEmu = -BULLET_INDENT_EMU;
			props.bullet = { kind: 'char', char: '•' };
		}
		return Object.keys(props).length ? props : undefined;
	};

	const finish = (): void => {
		current = null;
	};
	const ensure = (): DrawingParagraph | null => {
		if (current) return current;
		if (paragraphs.length >= bounds.maxParagraphs) {
			truncated = stopped = true;
			return null;
		}
		const properties = paragraphProperties();
		current = { items: [], ...(properties ? { properties } : {}) };
		paragraphs.push(current);
		return current;
	};
	const lastChar = (paragraph: DrawingParagraph): string => {
		const last = paragraph.items[paragraph.items.length - 1];
		if (!last) return '';
		return last.kind === 'run' ? last.text.slice(-1) : '\n';
	};
	const addText = (raw: string, verbatim = false): void => {
		const frame = top();
		if (marker !== null || frame.marker) {
			if (marker !== null) marker += raw;
			return;
		}
		let text =
			frame.pre || verbatim ? raw.replace(/\r\n?/g, '\n') : raw.replace(/[ \t\n\r\f]+/g, ' ');
		if (!frame.pre && !current && text.trim() === '') return;
		const paragraph = ensure();
		if (!paragraph) return;
		if (!frame.pre && text.startsWith(' ') && /^$|[ \n]/.test(lastChar(paragraph)))
			text = text.slice(1);
		text = text.replace(/ /g, ' ');
		if (textLength + text.length > bounds.maxTextLength) {
			text = text.slice(0, Math.max(0, bounds.maxTextLength - textLength));
			truncated = stopped = true;
		}
		textLength += text.length;
		const lines = text.split('\n');
		for (const [index, line] of lines.entries()) {
			if (index > 0) paragraph.items.push({ kind: 'break', properties: { ...frame.run } });
			if (!line) continue;
			const last = paragraph.items[paragraph.items.length - 1];
			if (last?.kind === 'run' && sameProps(last.properties, frame.run)) last.text += line;
			else
				paragraph.items.push({
					kind: 'run',
					text: line,
					...(Object.keys(frame.run).length ? { properties: { ...frame.run } } : {}),
				});
		}
	};

	/** Closes the innermost `name` (and what it holds); false when none is open. */
	const close = (name: string): boolean => {
		for (let i = stack.length - 1; i > 0; i--) {
			if ((stack[i] as Frame).name !== name) continue;
			while (stack.length > i) {
				const frame = stack.pop() as Frame;
				if (BLOCKS.has(frame.name)) finish();
			}
			return true;
		}
		return false;
	};

	for (const token of tokens) {
		if (stopped) break;
		if (token.kind === 'text') {
			if (token.text) addText(token.text);
			continue;
		}
		if (token.kind === 'conditional') {
			// `<![if !supportLists]>1.<![endif]>`: Word's typed marker, read as the paragraph's bullet.
			if (token.open && token.supportLists) {
				conditionalMarker = true;
				marker = '';
			} else if (!token.open && conditionalMarker) {
				conditionalMarker = false;
				const typed = marker ?? '';
				marker = null;
				applyMarker(typed);
			}
			continue;
		}
		if (token.kind === 'close') {
			const frame = top();
			if (frame.marker && frame.name === token.name) {
				stack.pop();
				const typed = marker ?? '';
				marker = null;
				applyMarker(typed);
				continue;
			}
			close(token.name);
			continue;
		}
		const { name, attrs } = token;
		if (name === 'br') {
			const paragraph = ensure();
			paragraph?.items.push({ kind: 'break', properties: { ...top().run } });
			continue;
		}
		if (name === 'hr') {
			finish();
			continue;
		}
		if ((name === 'td' || name === 'th') && current) {
			addText('\t', true);
		}
		if (VOID.has(name) || token.selfClosing) continue;
		if (name === 'p' || name === 'li' || /^h[1-6]$/.test(name) || name === 'div') {
			if (top().name === 'p') close('p');
		}
		if (name === 'li') {
			for (let i = stack.length - 1; i > 0; i--) {
				const open = (stack[i] as Frame).name;
				if (open === 'ul' || open === 'ol') break;
				if (open === 'li') {
					close('li');
					break;
				}
			}
		}
		if (BLOCKS.has(name)) finish();
		if (stack.length > bounds.maxDepth) {
			truncated = true;
			continue;
		}
		const parent = top();
		const block = BLOCKS.has(name);
		const run: DrawingRunProperties = {
			...parent.run,
			...tagRunProperties(name, attrs),
			...cssRunProperties(attrs.get('style') ?? '', !block),
		};
		const href = name === 'a' ? attrs.get('href')?.trim() : undefined;
		if (href && !href.startsWith('#') && hyperlinkPolicy(href).store) {
			const relId = `htmlLink${Object.keys(hyperlinks).length + 1}`;
			hyperlinks[relId] = href;
			const tooltip = attrs.get('title')?.slice(0, 256);
			run.hyperlinkClick = { relId, ...(tooltip ? { tooltip } : {}) };
		}
		const isMarker = /mso-list\s*:\s*ignore/i.test(attrs.get('style') ?? '');
		const frame: Frame = {
			name,
			run,
			pre: parent.pre || name === 'pre',
			marker: isMarker,
		};
		if (block) {
			const align = blockAlignment(name, attrs);
			if (align) frame.align = align;
			const level = name === 'p' || name === 'div' ? msoListLevel(attrs) : undefined;
			if (level !== undefined) frame.msoLevel = level;
		}
		if (name === 'ul' || name === 'ol') {
			frame.ordered = name === 'ol';
			frame.attrs = attrs;
		}
		stack.push(frame);
		if (isMarker && marker === null) marker = '';
	}

	function applyMarker(typed: string): void {
		const paragraph = ensure();
		if (!paragraph || !typed.trim()) return;
		paragraph.properties = { ...paragraph.properties, bullet: bulletFromMarker(typed) };
	}

	for (const paragraph of paragraphs) {
		const last = paragraph.items[paragraph.items.length - 1];
		if (last?.kind === 'run') {
			last.text = last.text.replace(/ +$/, '');
			if (!last.text) paragraph.items.pop();
		}
	}
	const empty = (p: DrawingParagraph | undefined) =>
		p !== undefined && p.items.length === 0 && !p.properties?.bullet;
	while (empty(paragraphs[paragraphs.length - 1])) paragraphs.pop();
	while (empty(paragraphs[0])) paragraphs.shift();
	return { paragraphs, hyperlinks, truncated };
}

/** The plain text of converted paragraphs, one line per paragraph. */
export function drawingParagraphsText(paragraphs: readonly DrawingParagraph[]): string {
	return paragraphs
		.map((p) =>
			p.items
				.map((item) => (item.kind === 'break' ? '\n' : 'text' in item ? item.text : ''))
				.join(''),
		)
		.join('\n');
}
