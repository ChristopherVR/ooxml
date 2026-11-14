/**
 * Rich text paste into a PowerPoint text body: clipboard HTML (converted by the core's
 * format-neutral `htmlToDrawingParagraphs`) as `TextSegment`s, merged with the destination's
 * formatting and spliced into the body at the selection. Pure; the DOM side is
 * `inline-rich-paste.ts`.
 *
 * Formatting follows PowerPoint's paste options for text: Keep Source Formatting (the default)
 * keeps the pasted run formatting over the destination run's, Use Destination Theme drops the
 * pasted fonts and colours so the theme supplies them, Keep Text Only keeps the destination
 * formatting and the line structure alone.
 *
 * @module render/html-text-paste
 */
import { htmlToDrawingParagraphs } from 'ooxml-core/drawingml';
import type {
	DrawingParagraph,
	DrawingParagraphProperties,
	DrawingRunProperties,
	HtmlTextPasteLimits,
} from 'ooxml-core/drawingml';
import type { BulletInfo, TextSegment, TextStyle } from 'ooxml-core/pptx';

import { isParagraphSeparatorSegment } from './text-segment-paragraph-break';

/** The text paste choices PowerPoint offers (Picture applies to shapes only). */
export type TextPasteFormat = 'keep-source-formatting' | 'use-destination-theme' | 'keep-text-only';

const EMU_PER_PX = 9525;
const ALIGN: Record<string, TextStyle['align']> = {
	l: 'left',
	ctr: 'center',
	r: 'right',
	just: 'justify',
};

const hex = (value: string | undefined): string | undefined =>
	value ? `#${value.toUpperCase()}` : undefined;

/** DrawingML run properties as a PowerPoint run style. */
function runStyle(
	props: DrawingRunProperties | undefined,
	links: Record<string, string>,
): TextStyle {
	const style: TextStyle = {};
	if (!props) {
		return style;
	}
	if (props.bold !== undefined) style.bold = props.bold;
	if (props.italic !== undefined) style.italic = props.italic;
	if (props.underline !== undefined) style.underline = props.underline !== 'none';
	if (props.strike !== undefined) style.strikethrough = props.strike !== 'noStrike';
	if (props.baseline !== undefined) style.baseline = Math.round(props.baseline * 100000);
	if (props.sizePt !== undefined) style.fontSize = props.sizePt;
	if (props.latin?.typeface) style.fontFamily = props.latin.typeface;
	if (props.caps === 'all') style.textCaps = 'all';
	if (props.caps === 'small') style.textCaps = 'small';
	if (props.fill?.kind === 'solid' && props.fill.color.kind === 'srgb') {
		style.color = hex(props.fill.color.value);
	}
	if (props.highlight?.kind === 'srgb') style.highlightColor = hex(props.highlight.value);
	const link = props.hyperlinkClick?.relId ? links[props.hyperlinkClick.relId] : undefined;
	if (link) {
		style.hyperlink = link;
		if (props.hyperlinkClick?.tooltip) style.hyperlinkTooltip = props.hyperlinkClick.tooltip;
	}
	return style;
}

function bulletInfo(props: DrawingParagraphProperties | undefined): BulletInfo | undefined {
	const bullet = props?.bullet;
	if (bullet?.kind === 'char') return { char: bullet.char };
	if (bullet?.kind === 'autoNumber') {
		return {
			autoNumType: bullet.scheme,
			...(bullet.startAt !== undefined ? { autoNumStartAt: bullet.startAt } : {}),
		};
	}
	return undefined;
}

/** Paragraph metadata a pasted paragraph's first segment carries. */
function paragraphMetadata(props: DrawingParagraphProperties | undefined): Partial<TextSegment> {
	const meta: Partial<TextSegment> = {};
	const bullet = bulletInfo(props);
	if (bullet) meta.bulletInfo = bullet;
	if (props?.level) meta.paragraphLevel = props.level;
	const paragraph: TextStyle = {};
	const align = props?.align ? ALIGN[props.align] : undefined;
	if (align) paragraph.align = align;
	if (props?.marginLeftEmu !== undefined) {
		paragraph.paragraphMarginLeft = props.marginLeftEmu / EMU_PER_PX;
	}
	if (props?.indentEmu !== undefined) paragraph.paragraphIndent = props.indentEmu / EMU_PER_PX;
	if (Object.keys(paragraph).length) meta.paragraphProperties = paragraph;
	return meta;
}

/**
 * Sets `bulletInfo.paragraphIndex` (the runtime ordinal offset the renderer adds to the start
 * value) on pasted numbered paragraphs, as the loader does: consecutive items at one level with
 * one scheme count up; an unnumbered paragraph ends the lists at its level and deeper, and a
 * numbered one ends the lists nested beneath it.
 */
function createNumberingCounter(): (meta: Partial<TextSegment>) => void {
	const running = new Map<number, { scheme: string; index: number }>();
	const forgetFrom = (level: number) => {
		for (const key of [...running.keys()]) if (key >= level) running.delete(key);
	};
	return (meta) => {
		const level = meta.paragraphLevel ?? 0;
		const scheme = meta.bulletInfo?.autoNumType;
		if (!scheme || !meta.bulletInfo) {
			forgetFrom(level);
			return;
		}
		forgetFrom(level + 1);
		const previous = running.get(level);
		const index = previous?.scheme === scheme ? previous.index + 1 : 0;
		running.set(level, { scheme, index });
		meta.bulletInfo = { ...meta.bulletInfo, paragraphIndex: index };
	};
}

/** Hyperlink and colour-choice fields of a destination run that a pasted value replaces. */
function mergeStyle(destination: TextStyle, source: TextStyle): TextStyle {
	const merged: TextStyle = { ...destination };
	for (const key of ['hyperlink', 'hyperlinkRId', 'hyperlinkTooltip', 'hyperlinkAction'] as const) {
		delete merged[key];
	}
	delete merged.hyperlinkTargetSlideIndex;
	if (source.color) {
		delete merged.colorRef;
		delete merged.colorXml;
	}
	if (source.highlightColor) delete merged.highlightColorXml;
	return { ...merged, ...source };
}

/** Whether converted paragraphs carry anything a plain-text paste would lose. */
export function hasRichPasteContent(paragraphs: readonly DrawingParagraph[]): boolean {
	return paragraphs.some(
		(p) =>
			Boolean(p.properties && Object.keys(p.properties).length) ||
			p.items.some(
				(item) => 'properties' in item && item.properties && Object.keys(item.properties).length,
			),
	);
}

/**
 * Pasted clipboard HTML as segments in `format`, every run over `destination` (the style at the
 * caret). Paragraphs after the first carry their own bullets, levels and alignment; the first
 * merges into the destination paragraph. `undefined` when the HTML holds no text.
 */
export function htmlToPasteSegments(
	html: string,
	destination: TextStyle,
	format: TextPasteFormat = 'keep-source-formatting',
	limits?: Partial<HtmlTextPasteLimits>,
): { segments: TextSegment[]; text: string; rich: boolean } | undefined {
	const { paragraphs, hyperlinks } = htmlToDrawingParagraphs(html, limits);
	if (!paragraphs.some((p) => p.items.length)) {
		return undefined;
	}
	const segments: TextSegment[] = [];
	const numbering = createNumberingCounter();
	for (const [index, paragraph] of paragraphs.entries()) {
		const meta = format === 'keep-text-only' ? {} : paragraphMetadata(paragraph.properties);
		numbering(meta);
		if (index > 0) {
			segments.push({ text: '\n', style: { ...destination }, isParagraphBreak: true });
		}
		const start = segments.length;
		for (const item of paragraph.items) {
			if (item.kind !== 'run' && item.kind !== 'break') continue;
			let source = format === 'keep-text-only' ? {} : runStyle(item.properties, hyperlinks);
			if (format === 'use-destination-theme') {
				const { color: _color, fontFamily: _font, ...rest } = source;
				source = rest;
			}
			const style = mergeStyle(destination, source);
			segments.push(
				item.kind === 'break'
					? { text: '\n', style, isLineBreak: true }
					: { text: item.text, style },
			);
		}
		if (segments.length === start) {
			segments.push({ text: '', style: { ...destination } });
		}
		if (index > 0) Object.assign(segments[start] as TextSegment, meta);
	}
	const text = segments.map((segment) => segment.text).join('');
	return { segments, text, rich: format !== 'keep-text-only' && hasRichPasteContent(paragraphs) };
}

const PARAGRAPH_KEYS = [
	'bulletInfo',
	'paragraphLevel',
	'paragraphProperties',
	'endParaRunProperties',
	'paragraphInsertionStyle',
	'emptyParagraphPropertiesAuthored',
	'bareParagraph',
] as const;

function withoutParagraphMetadata(segment: TextSegment): TextSegment {
	const copy = { ...segment, style: { ...segment.style } };
	for (const key of PARAGRAPH_KEYS) delete copy[key];
	return copy;
}

const lengthOf = (segment: TextSegment): number =>
	segment.isLineBreak || isParagraphSeparatorSegment(segment) ? 1 : segment.text.length;

/** Splits `segments` at body offset `offset` (a break counts as one character). */
function splitAt(segments: readonly TextSegment[], offset: number): [TextSegment[], TextSegment[]] {
	let position = 0;
	for (const [index, segment] of segments.entries()) {
		const length = lengthOf(segment);
		if (offset <= position) {
			return [segments.slice(0, index), segments.slice(index)];
		}
		if (
			offset < position + length &&
			!segment.isLineBreak &&
			!isParagraphSeparatorSegment(segment)
		) {
			const cut = offset - position;
			const head = { ...segment, text: segment.text.slice(0, cut) };
			const tail = { ...withoutParagraphMetadata(segment), text: segment.text.slice(cut) };
			return [
				[...segments.slice(0, index), head],
				[tail, ...segments.slice(index + 1)],
			];
		}
		position += length;
	}
	return [[...segments], []];
}

/** The run style at body offset `offset`: the run the caret follows, else the one it precedes. */
export function styleAtOffset(
	segments: readonly TextSegment[],
	offset: number,
	fallback: TextStyle = {},
): TextStyle {
	const [before, after] = splitAt(segments, offset);
	const isText = (segment: TextSegment) =>
		!segment.isLineBreak && !isParagraphSeparatorSegment(segment) && segment.text.length > 0;
	const previous = [...before].reverse().find(isText);
	const style = previous?.style ?? after.find(isText)?.style ?? fallback;
	return { ...style };
}

/**
 * Replaces body offsets `start..end` with `inserted`. The first pasted paragraph joins the
 * paragraph the selection starts in and keeps that paragraph's properties.
 */
export function spliceTextSegments(
	segments: readonly TextSegment[],
	start: number,
	end: number,
	inserted: readonly TextSegment[],
): TextSegment[] {
	const [before] = splitAt(segments, start);
	const [, after] = splitAt(segments, Math.max(start, end));
	const pasted = inserted.map((segment) => ({ ...segment, style: { ...segment.style } }));
	const last = before[before.length - 1];
	const atParagraphStart = !last || isParagraphSeparatorSegment(last);
	const tail = after.map((segment) => ({ ...segment }));
	if (atParagraphStart && tail[0] && pasted[0] && !isParagraphSeparatorSegment(tail[0])) {
		// The destination paragraph's own properties move onto the pasted text that now begins it.
		for (const key of PARAGRAPH_KEYS) {
			if (tail[0][key] !== undefined) {
				(pasted[0] as unknown as Record<string, unknown>)[key] = tail[0][key];
				delete tail[0][key];
			}
		}
	}
	return [...before, ...pasted, ...tail].filter(
		(segment, index, all) =>
			segment.text !== '' || all.length === 1 || hasParagraphMetadata(segment),
	);
}

function hasParagraphMetadata(segment: TextSegment): boolean {
	return PARAGRAPH_KEYS.some((key) => segment[key] !== undefined);
}

/** The body text of segments, breaks as `\n`. */
export function segmentsBodyText(segments: readonly TextSegment[]): string {
	return segments
		.map((segment) =>
			segment.isLineBreak || isParagraphSeparatorSegment(segment) ? '\n' : segment.text,
		)
		.join('');
}
