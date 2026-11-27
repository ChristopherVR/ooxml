/**
 * The portfolio's design system: one palette, one type scale and one grid, plus the few text
 * helpers every slide uses so that kickers, captions and figures look the same everywhere.
 */
import type { PptxElement } from '../../../../src/core/pptx/core/types/elements';
import { FONTS } from '../../lib/fiction';
import { pt } from '../kit/geometry';
import type { Paragraph, Slide, TextBoxOptions } from '../kit/slide';

/** Bone paper, charcoal ink and a single vermilion accent; the rest stays quiet. */
export const C = {
	bone: '#F6F3EE',
	paper: '#ECE6DC',
	ink: '#1D1D1B',
	soft: '#6F6A63',
	rule: '#CFC7BA',
	vermilion: '#E2452B',
	clay: '#D8C9B8',
	sage: '#9AA791',
	moss: '#56634F',
	blue: '#2B3A67',
	wheat: '#D6A24A',
	night: '#161A26',
} as const;

/** Type sizes in points. */
export const T = {
	display: 64,
	title: 34,
	heading: 26,
	lead: 19,
	body: 12.5,
	caption: 10,
	kicker: 8.5,
	figure: 40,
} as const;

/** The grid: 80 px side margins, twelve 71 px columns with 24 px gutters. */
export const G = {
	left: 80,
	right: 1200,
	width: 1120,
	top: 64,
	bottom: 616,
	column: 71.33,
	gutter: 24,
	footerRule: 648,
	footerText: 662,
} as const;

/** The left edge of a grid column (0 to 11). */
export function col(index: number): number {
	return G.left + index * (G.column + G.gutter);
}

/** The width of a span of grid columns. */
export function span(columns: number): number {
	return columns * G.column + (columns - 1) * G.gutter;
}

const FLUSH = { insets: { left: 0, top: 0, right: 0, bottom: 0 }, autoFit: 'none' } as const;

/** A text box with no insets, so its text sits exactly on the grid. */
export function label(
	slide: Slide,
	name: string,
	content: string | Paragraph[],
	options: TextBoxOptions,
): PptxElement {
	return slide.text(name, content, { ...FLUSH, wrap: true, color: C.ink, ...options });
}

/** A small tracked, upper-case line that names what follows. */
export function kicker(
	slide: Slide,
	name: string,
	text: string,
	x: number,
	y: number,
	width: number,
	color: string = C.soft,
): PptxElement {
	return label(
		slide,
		name,
		[[{ text: text.toUpperCase(), style: { characterSpacing: 1.6, bold: true } }]],
		{ x, y, width, height: 16, fontSize: pt(T.kicker), color },
	);
}

/** A display line in the heading face. */
export function headline(
	slide: Slide,
	name: string,
	text: string | Paragraph[],
	options: TextBoxOptions & { size: number },
): PptxElement {
	const { size, ...rest } = options;
	const tracking = -size / 55;
	const paragraphs: Paragraph[] =
		typeof text === 'string'
			? text.split('\n').map((line) => [{ text: line, style: { characterSpacing: tracking } }])
			: text;
	return label(slide, name, paragraphs, {
		fontFamily: FONTS.display,
		fontSize: pt(size),
		bold: true,
		lineSpacing: 0.95,
		...rest,
	});
}

/** A hairline rule. */
export function rule(
	slide: Slide,
	name: string,
	x: number,
	y: number,
	width: number,
	color: string = C.rule,
): PptxElement {
	return slide.shape(name, 'rect', {
		x,
		y,
		width,
		height: 1,
		fill: { type: 'solid', color },
		stroke: { color: 'transparent' },
	});
}

/** A result figure: a large number over a caption. Returns the number and the caption. */
export function figure(
	slide: Slide,
	name: string,
	x: number,
	y: number,
	width: number,
	value: string,
	caption: string,
	color: string = C.ink,
): [PptxElement, PptxElement] {
	const number = headline(slide, `${name} value`, value, {
		x,
		y,
		width,
		height: 52,
		size: T.figure,
		color,
	});
	const text = label(slide, `${name} caption`, caption, {
		x,
		y: y + 58,
		width,
		height: 34,
		fontSize: pt(T.caption),
		color: color === C.ink ? C.soft : color,
		lineSpacing: 1.25,
	});
	return [number, text];
}

/** The slide number field in the footer (the layouts keep its place at idx 12). */
export function numbered(slide: Slide): void {
	slide.placeholder('sldNum', String(slide.number).padStart(2, '0'), 12);
}
