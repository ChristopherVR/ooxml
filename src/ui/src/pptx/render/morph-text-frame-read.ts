/**
 * One half of a carried-text morph: where its text is anchored and how wide
 * each of its lines is. See `morph-text-frame` for what PowerPoint does with it.
 *
 * @module render/morph-text-frame-read
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { getSubstituteFontFamily, hasShapeProperties, hasTextProperties } from 'ooxml-core/pptx';

import {
	DEFAULT_BODY_INSET_LR_PX,
	DEFAULT_BODY_INSET_TB_PX,
	DEFAULT_FONT_FAMILY,
	DEFAULT_TEXT_FONT_SIZE,
} from '../constants';
import type { RenderParagraph } from './paragraph-types';
import { resolveTextBodyRectPadding } from './text-body-rect';
import { buildParagraphs } from './text-paragraphs';
import { resolveLineHeight } from './text-line-height';
import { resolveCssTextAlign } from './text-paragraph-style';
import { resolveAutoFitFontScale } from './text-style-helpers';

/**
 * Advance width of `text` painted in the canvas `font` shorthand, or
 * `undefined` when nothing can measure (no DOM).
 */
export type MorphTextMeasure = (text: string, font: string, kerning: boolean) => number | undefined;

let canvasContext: CanvasRenderingContext2D | null | undefined;

/** The default measurer: a 2D canvas, the same one the text renderer uses. */
export const measureMorphTextWithCanvas: MorphTextMeasure = (text, font, kerning) => {
	if (canvasContext === undefined) {
		canvasContext =
			typeof document === 'undefined' ? null : document.createElement('canvas').getContext('2d');
	}
	if (!canvasContext) {
		return undefined;
	}
	canvasContext.font = font;
	canvasContext.fontKerning = kerning ? 'auto' : 'none';
	const width = canvasContext.measureText(text).width;
	return Number.isFinite(width) ? width : undefined;
};

/** What one half contributes: its anchor and its paragraphs' line extents. */
export interface TextFrame {
	anchorX: number;
	anchorY: number;
	align: string;
	vAlign: string;
	/** Column width the paragraphs lay out in. */
	available: number;
	wraps: boolean;
	/** Left and top body insets, in px from the element's own box. */
	insetLeft: number;
	insetTop: number;
	/** Bottom body inset and the element's height, for middle and bottom anchors. */
	insetBottom: number;
	boxHeight: number;
	/** `a:bodyPr/@anchorCtr`: the paragraphs are centred as one shrink-wrapped block. */
	anchorCenter: boolean;
	paragraphs: FrameParagraph[];
}

/** One paragraph of a {@link TextFrame}. */
export interface FrameParagraph {
	text: string;
	fonts: string[];
	/** Largest run font size (px), the paragraph's line extent for alignment. */
	size: number;
	runs: RenderParagraph['runs'];
	/** Font size (px) of the paragraph's own strut: its line boxes are at least this tall. */
	strut: number;
	/** The canvas font of the paragraph's strut (the body's face at its strut size). */
	strutFont: string;
	/** The CSS `line-height` the paragraph lays out with. */
	lineHeight: number | string;
	/** `margin-bottom` (px) after the paragraph. */
	spaceAfterPx: number;
	/** `margin-left` and `text-indent` (px) the paragraph is laid out with. */
	marginLeftPx: number;
	textIndentPx: number;
	/** Font size (px) of the paragraph element itself, before its runs. */
	ownStrut: number;
	/** The paragraph's own `text-align` and `text-align-last`, as painted. */
	textAlign: string;
	textAlignLast: string | undefined;
}

/** Whether `element` is (or holds) a picture or a group, never carried by its text. */
function paintsPicture(element: PptxElement): boolean {
	const image = element as { imagePath?: string; svgPath?: string; children?: unknown[] };
	return Boolean(image.imagePath || image.svgPath || image.children);
}

/**
 * Whether `element` paints anything besides its text: a picture, a group, a
 * fill or an outline. Such a pair's text can only be carried apart from its
 * body (see `morph-text-body`).
 */
export function paintsBody(element: PptxElement): boolean {
	if (paintsPicture(element)) {
		return true;
	}
	if (!hasShapeProperties(element) || !element.shapeStyle) {
		return false;
	}
	const style = element.shapeStyle;
	const filled =
		style.fillMode !== 'none' &&
		(style.fillOpacity ?? 1) > 0 &&
		Boolean(
			style.fillColor ||
			style.fillGradient ||
			style.fillGradientStops?.length ||
			style.fillPatternPreset ||
			style.fillImageUrl,
		);
	const stroked = Boolean(style.strokeColor) && (style.strokeWidth ?? 0) > 0;
	return filled || stroked;
}

function pxOf(value: string | number | undefined): number | undefined {
	if (typeof value === 'number') {
		return value;
	}
	const parsed = typeof value === 'string' ? Number.parseFloat(value) : Number.NaN;
	return Number.isFinite(parsed) ? parsed : undefined;
}

/** The canvas font a run paints with: its own declarations, else the body's. */
function runFont(
	style: Record<string, string | number>,
	block: { family: string; size: number },
): string {
	const size = pxOf(style.fontSize) ?? block.size;
	const family = (style.fontFamily as string | undefined) ?? block.family;
	const weight = style.fontWeight ?? 'normal';
	const italic = style.fontStyle === 'italic' ? 'italic ' : '';
	return `${italic}${weight} ${size}px ${family}`;
}

/** What {@link readFrame} may accept beyond a plain, upright, body-less text box. */
export interface ReadFrameOptions {
	/**
	 * Accept a rotated box. Its frame is read in the box's own, unrotated
	 * coordinates, so the caller has to rotate whatever it derives from it.
	 */
	rotated?: boolean;
	/** Accept a box that paints a fill or an outline, reading only its text. */
	withBody?: boolean;
}

export function readFrame(
	element: PptxElement,
	options: ReadFrameOptions = {},
): TextFrame | undefined {
	if (
		!hasTextProperties(element) ||
		!element.text ||
		(!options.rotated && (element.rotation ?? 0) !== 0) ||
		element.flipHorizontal ||
		element.flipVertical ||
		(options.withBody ? paintsPicture(element) : paintsBody(element))
	) {
		return undefined;
	}
	const ts = element.textStyle;
	if (
		(ts?.textDirection && ts.textDirection !== 'horizontal') ||
		(ts?.columnCount ?? 1) > 1 ||
		(ts?.textWarpPreset && ts.textWarpPreset !== 'textNoShape') ||
		ts?.rtl
	) {
		return undefined;
	}
	const fontScale = resolveAutoFitFontScale(ts);
	const block = {
		family: ts?.fontFamily ? getSubstituteFontFamily(ts.fontFamily) : DEFAULT_FONT_FAMILY,
		size: (ts?.fontSize || DEFAULT_TEXT_FONT_SIZE) * fontScale,
	};
	// The geometry's own text rectangle insets the body further, as the
	// renderer lays it out (`text-body-rect`): a rounded rectangle's text sits
	// inside its corners.
	const rect = resolveTextBodyRectPadding(element);
	const left = (ts?.bodyInsetLeft ?? DEFAULT_BODY_INSET_LR_PX) + rect.left;
	const right = (ts?.bodyInsetRight ?? DEFAULT_BODY_INSET_LR_PX) + rect.right;
	const top = (ts?.bodyInsetTop ?? DEFAULT_BODY_INSET_TB_PX) + rect.top;
	const bottom = (ts?.bodyInsetBottom ?? DEFAULT_BODY_INSET_TB_PX) + rect.bottom;
	const bodyAlign = resolveCssTextAlign(ts?.align, false);
	const bodyLineHeight = resolveLineHeight(ts, false);
	let align: string | undefined;
	const paragraphs: TextFrame['paragraphs'] = [];
	for (const paragraph of buildParagraphs(element)) {
		if (
			paragraph.bulletMarker ||
			paragraph.bulletPicture ||
			paragraph.rtl ||
			paragraph.runs.some((run) => run.tabLines || run.equation || run.ruby)
		) {
			return undefined;
		}
		const text = paragraph.runs.map((run) => run.text).join('');
		const fonts = paragraph.runs.map((run) => runFont(run.style, block));
		const sizes = paragraph.runs.map((run) => pxOf(run.style.fontSize) ?? block.size);
		const size = Math.max(paragraph.strutFontSizePx ?? 0, ...sizes, 0) || block.size;
		paragraphs.push({
			text,
			fonts,
			size,
			runs: paragraph.runs,
			strut: Math.max(paragraph.strutFontSizePx ?? block.size, ...sizes, 0),
			strutFont: runFont(
				{},
				{ family: block.family, size: paragraph.strutFontSizePx ?? block.size },
			),
			lineHeight: paragraph.lineHeight ?? bodyLineHeight,
			spaceAfterPx: paragraph.spaceAfterPx ?? 0,
			marginLeftPx: paragraph.marginLeftPx ?? 0,
			textIndentPx: paragraph.textIndentPx ?? 0,
			ownStrut: paragraph.strutFontSizePx ?? block.size,
			textAlign: (paragraph.paragraphStyle?.textAlign as string | undefined) ?? bodyAlign ?? 'left',
			textAlignLast: paragraph.paragraphStyle?.textAlignLast as string | undefined,
		});
		if (text.trim() === '') {
			continue;
		}
		const own = (paragraph.paragraphStyle?.textAlign as string | undefined) ?? bodyAlign ?? 'left';
		// A justified single line is laid out like a left-aligned one; `dist`
		// spreads even its last line and has no fixed anchor.
		const effective = own === 'justify' && ts?.align !== 'dist' ? 'left' : own;
		if (align !== undefined && align !== effective) {
			return undefined;
		}
		align = effective;
	}
	if (align === undefined || !['left', 'center', 'right'].includes(align)) {
		return undefined;
	}
	const available = element.width - left - right;
	const vAlign =
		ts?.vAlign === 'bottom' ? 'bottom' : ts?.vAlign && ts.vAlign !== 'top' ? 'middle' : 'top';
	return {
		anchorX:
			align === 'left'
				? element.x + left
				: align === 'right'
					? element.x + element.width - right
					: element.x + left + available / 2,
		anchorY:
			vAlign === 'top'
				? element.y + top
				: vAlign === 'bottom'
					? element.y + element.height - bottom
					: element.y + top + (element.height - top - bottom) / 2,
		align,
		vAlign,
		available,
		wraps: ts?.textWrap !== 'none',
		insetLeft: left,
		insetTop: top,
		insetBottom: bottom,
		boxHeight: element.height,
		anchorCenter: ts?.anchorCenter === true,
		paragraphs,
	};
}

export function paragraphWidth(
	paragraph: FrameParagraph,
	measure: MorphTextMeasure,
): number | undefined {
	let width = 0;
	for (const [index, run] of paragraph.runs.entries()) {
		const font = paragraph.fonts[index];
		const advance = run.text ? measure(run.text, font ?? '', run.style.fontKerning !== 'none') : 0;
		if (advance === undefined) {
			return undefined;
		}
		width += advance + (pxOf(run.style.letterSpacing) ?? 0) * [...run.text].length;
	}
	return width;
}
