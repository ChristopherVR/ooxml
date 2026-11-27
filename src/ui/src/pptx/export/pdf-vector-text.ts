/**
 * Read the text of a mounted slide as line fragments: each run of words that the browser laid
 * out on one line, with its measured position and width (`Range.getClientRects`). Called while
 * the stage's transforms are neutralised, so the boxes are the untransformed layout and the
 * element's matrix is applied when painting.
 */
import type { PdfFontSet } from 'ooxml-core/pdf/write';
import type { PdfExportDiagnosticCode } from './pdf-export-diagnostics';
import { cssColor } from './pdf-vector-css';
import { canWritePdfText } from './pdf-vector-paint';
import type { Matrix, PdfClip, TextItem } from './pdf-vector-display-list';

/** Scripts whose glyphs need shaping or bidi reordering a PDF text run cannot express. */
const COMPLEX_SCRIPT = new RegExp(
	'[\u0590-\u08ff\u0900-\u0dff\u0f00-\u0fff\u1000-\u109f\u1780-\u17ff\u1800-\u18af\ua840-\ua87f\ufb1d-\ufdff\ufe70-\ufeff\u200f\u202b\u202e]|\u200d|\\p{M}',
	'u',
);

/** Whether `text` needs shaping (joining scripts, combining marks, right-to-left). */
export const needsShaping = (text: string): boolean => COMPLEX_SCRIPT.test(text);

/** The most words read per slide before the whole slide is drawn as a picture. */
export const MAX_TEXT_TOKENS = 20_000;

interface Token {
	start: number;
	end: number;
	left: number;
	right: number;
	top: number;
	height: number;
}

function rectsOf(range: Range): DOMRect[] {
	return Array.from(range.getClientRects()).filter((rect) => rect.width > 0 && rect.height > 0);
}

function applyTextTransform(text: string, transform: string): string {
	switch (transform) {
		case 'uppercase':
			return text.toUpperCase();
		case 'lowercase':
			return text.toLowerCase();
		case 'capitalize':
			return text.replace(
				/(^|\s)(\S)/g,
				(_, space: string, ch: string) => space + ch.toUpperCase(),
			);
		default:
			return text;
	}
}

/** What measuring one text node found. */
export interface TextNodeResult {
	items: TextItem[];
	tokens: number;
	/** Why the node needs a picture instead (unset when it can be text). */
	raster?: PdfExportDiagnosticCode;
}

export interface TextNodeContext {
	matrix: Matrix;
	clips: readonly PdfClip[];
	alpha: number;
	/** Offset of the stage's top-left corner in viewport px. */
	originX: number;
	originY: number;
	/** Viewport px per stage CSS px. */
	scale: number;
	fonts?: PdfFontSet;
	invisible?: boolean;
}

/** Ascent share of the content box for a font, from an embedded font's metrics when known. */
function ascentRatio(family: string, bold: boolean, italic: boolean, fonts?: PdfFontSet): number {
	const font = fonts?.has(family) ? fonts.pick(family, bold, italic, 'x') : undefined;
	if (font) {
		const ascent = font.font.ascent,
			descent = Math.abs(font.font.descent);
		if (ascent > 0 && ascent + descent > 0) return ascent / (ascent + descent);
	}
	// Typical sans-serif metrics (ascent 0.905, descent 0.212 em).
	return 0.81;
}

/** Measure one text node into fragments. */
export function measureTextNode(
	node: Text,
	style: CSSStyleDeclaration,
	context: TextNodeContext,
	tokenBudget: number,
): TextNodeResult {
	const raw = node.data;
	if (!raw.trim()) return { items: [], tokens: 0 };
	// Vertical text is a picture; its characters are measured one by one for the text layer.
	const vertical = !/^horizontal/.test(style.writingMode || 'horizontal-tb');
	const display = applyTextTransform(raw, style.textTransform);
	const shaping = needsShaping(display);
	const weight = style.fontWeight;
	const bold = weight === 'bold' || Number.parseInt(weight, 10) >= 600;
	const italic = /italic|oblique/.test(style.fontStyle);
	const family = style.fontFamily;
	// Characters no supplied or standard font has (CJK in a deck that embeds no font for it) are
	// drawn as a picture of the browser's rendering, with the text kept as an invisible layer.
	const noFont =
		!shaping &&
		!context.invisible &&
		!canWritePdfText(display.replace(/\s+/g, ' ').trim(), family, bold, italic, context.fonts);
	const picture = vertical || shaping || noFont;
	const color = cssColor(style.getPropertyValue('-webkit-text-fill-color') || style.color);
	if (color === undefined) return { items: [], tokens: 0, raster: 'textColourAsPicture' };
	const doc = node.ownerDocument;
	const range = doc.createRange();
	const pre = /^(pre|pre-wrap|break-spaces)$/.test(style.whiteSpace);
	const tokens: Token[] = [];
	const k = context.scale;
	const token = (start: number, end: number, r: DOMRect): Token => ({
		start,
		end,
		left: (r.left - context.originX) / k,
		right: (r.right - context.originX) / k,
		top: (r.top - context.originY) / k,
		height: r.height / k,
	});
	const words = [...raw.matchAll(/\S+/g)];
	if (words.length > tokenBudget) return { items: [], tokens: words.length };
	for (const word of words) {
		const start = word.index!;
		const end = start + word[0].length;
		range.setStart(node, start);
		range.setEnd(node, end);
		const rects = rectsOf(range);
		if (rects.length === 1 && (!vertical || end - start === 1)) {
			const r = rects[0]!;
			tokens.push(token(start, end, r));
			continue;
		}
		if (!rects.length) continue;
		// A word broken across lines: measure it character by character.
		for (let i = start; i < end; i++) {
			range.setStart(node, i);
			range.setEnd(node, i + 1);
			const r = rectsOf(range)[0];
			if (r) tokens.push(token(i, i + 1, r));
		}
	}
	range.detach?.();
	if (!tokens.length) return { items: [], tokens: words.length };
	const size = Number.parseFloat(style.fontSize) || 16;
	const ratio = ascentRatio(family, bold, italic, context.fonts);
	const decoration = style.textDecorationLine || '';
	const justify = /justify/.test(style.textAlign);
	const items: TextItem[] = [];
	const flush = (group: Token[]) => {
		const first = group[0]!,
			last = group[group.length - 1]!;
		const text = pre
			? display.slice(first.start, last.end)
			: group
					.map(
						(token, i) =>
							(i && token.start !== group[i - 1]!.end ? ' ' : '') +
							display.slice(token.start, token.end),
					)
					.join('');
		const height = Math.max(...group.map((token) => token.height));
		items.push({
			kind: 'text',
			matrix: context.matrix,
			clips: context.clips,
			alpha: context.alpha,
			text,
			x: first.left,
			baseline: first.top + height * ratio,
			width: last.right - first.left,
			size,
			family,
			bold,
			italic,
			color: color ?? { r: 0, g: 0, b: 0, a: 0 },
			...(decoration.includes('underline') ? { underline: true } : {}),
			...(decoration.includes('line-through') ? { strike: true } : {}),
			...(context.invisible || picture || !color ? { invisible: true } : {}),
		});
	};
	let group: Token[] = [];
	for (const token of tokens) {
		const previous = group[group.length - 1];
		const sameLine =
			previous &&
			!vertical &&
			Math.abs(token.top - previous.top) < Math.max(token.height, previous.height) * 0.5 &&
			token.left >= previous.right - 1 &&
			!justify &&
			// Characters of a broken word join without spaces; words join with one.
			(token.start === previous.end || display.slice(previous.end, token.start).trim() === '');
		if (previous && !sameLine) {
			flush(group);
			group = [];
		}
		group.push(token);
	}
	if (group.length) flush(group);
	return {
		items,
		tokens: words.length,
		...(picture && !context.invisible
			? {
					raster: vertical
						? ('verticalTextAsPicture' as const)
						: shaping
							? ('complexScriptAsPicture' as const)
							: ('noFontTextAsPicture' as const),
				}
			: {}),
	};
}
