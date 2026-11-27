/**
 * The design system of the deck: one palette, one type scale, one 12-column grid.
 *
 * Grid: 1280 x 720 slide, 64 px side margins, 12 columns of 74 px with 24 px gutters.
 * Rows: kicker at 44, headline from 68, content from 196 to 612, source line at 626, footer
 * rule at 660. Type scale (pt): 96 display, 54 cover, 28 headline, 18 lead, 14 body, 11 small,
 * 9.5 caption, 10 kicker (tracked capitals).
 */
import type { TextStyleInput } from '../../../../src/core/pptx/core/builders/sdk/index';
import type { PptxElement } from '../../../../src/core/pptx/core/types/elements';
import { FONTS } from '../../lib/fiction';
import { pt, type Box } from '../kit/geometry';
import type { Paragraph, Slide } from '../kit/slide';

export const C = {
	paper: '#FBF9F4',
	white: '#FFFFFF',
	ink: '#1B2430',
	slate: '#5B6878',
	teal: '#0E7C86',
	tealSoft: '#9CC9CC',
	coral: '#E4572E',
	sand: '#E9E2D0',
	sandDeep: '#CFC6B0',
	ochre: '#D9A441',
	mist: '#9AA5B4',
	rule: '#D8D2C4',
} as const;

/** Mode colours, the same on every chart and on the map. */
export const MODE_COLORS = [C.teal, C.coral, C.ochre, C.mist];

export const MARGIN = 64;
export const GUTTER = 24;
export const COLUMN = 74;
export const CONTENT_TOP = 196;
export const CONTENT_BOTTOM = 612;

/** The left edge and width of a span of grid columns (0-based start). */
export function span(start: number, count: number): { x: number; width: number } {
	return {
		x: MARGIN + start * (COLUMN + GUTTER),
		width: count * COLUMN + (count - 1) * GUTTER,
	};
}

export interface LabelStyle {
	/** Size in points. */
	size: number;
	color?: string;
	bold?: boolean;
	display?: boolean;
	align?: 'left' | 'center' | 'right';
	valign?: 'top' | 'middle' | 'bottom';
	/** Letter spacing in px. */
	tracking?: number;
	leading?: number;
	spaceBefore?: number;
}

function runStyle(style: LabelStyle): Partial<TextStyleInput> {
	return {
		fontSize: pt(style.size),
		fontFamily: style.display ? FONTS.display : FONTS.body,
		bold: style.bold ?? false,
		color: style.color ?? C.ink,
		...(style.tracking === undefined ? {} : { characterSpacing: style.tracking }),
	};
}

/** A text box that sits exactly on the grid: no insets, no autofit, one style. */
export function label(
	slide: Slide,
	name: string,
	content: string | Paragraph[],
	box: Box,
	style: LabelStyle,
): PptxElement {
	const base = runStyle(style);
	const paragraphs: Paragraph[] = (typeof content === 'string' ? [content] : content).map(
		(paragraph) =>
			typeof paragraph === 'string'
				? [{ text: paragraph, style: base }]
				: paragraph.map((run) => ({ text: run.text, style: { ...base, ...run.style } })),
	);
	return slide.text(name, paragraphs, {
		...box,
		fontSize: pt(style.size),
		fontFamily: style.display ? FONTS.display : FONTS.body,
		color: style.color ?? C.ink,
		alignment: style.align ?? 'left',
		verticalAlignment: style.valign ?? 'top',
		lineSpacing: style.leading ?? 1.15,
		insets: { left: 0, top: 0, right: 0, bottom: 0 },
		autoFit: 'none',
		...(style.spaceBefore === undefined ? {} : { spaceBefore: style.spaceBefore }),
	});
}

/** A hairline, as a filled 1 to 2 px rectangle (so it morphs and animates like any shape). */
export function rule(slide: Slide, name: string, box: Box, color: string = C.rule): PptxElement {
	return slide.shape(name, 'rect', {
		...box,
		fill: { type: 'solid', color },
		stroke: { color: 'transparent' },
	});
}

/** The fonts every chart uses for its small text. */
export const CHART_FONT = { family: FONTS.body, size: 10, color: C.slate };

/** Fills the layout's kicker, headline, source line and slide number. */
export function frame(
	slide: Slide,
	text: { kicker: string; headline: string; source?: string },
	number: number,
): void {
	slide.placeholder('title', text.headline);
	slide.placeholder('body', text.kicker, 10);
	if (text.source) slide.placeholder('body', text.source, 11);
	slide.placeholder('sldNum', String(number), 12);
}
