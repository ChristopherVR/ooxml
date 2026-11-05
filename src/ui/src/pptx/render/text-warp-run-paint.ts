import type { TextSegment, TextStyle } from 'ooxml-core/pptx';
import { getSubstituteFontFamily } from 'ooxml-core/pptx';

import { DEFAULT_FONT_FAMILY } from '../constants';
import { normalizeHexColor } from './fill-style';
import { fontWeightAttribute, resolveFontItalic } from './font-weight-names';

/** The paint of one warped run, as SVG presentation values every binding maps onto its template. */
export interface WarpRunPaint {
	/** SVG `fill`: an SVG `<text>` is painted by `fill`, never by CSS `color`. */
	fill: string;
	fontFamily: string;
	fontSizePx: number;
	/** `bold`, a numeric weight, or undefined for 400. */
	fontWeight: string | undefined;
	fontStyle: 'italic' | undefined;
	textDecoration: string | undefined;
}

const HYPERLINK_COLOR = '#0563C1';

/** Resolve a segment's SVG paint, falling back to the element's text style. */
export function resolveWarpRunPaint(
	segment: TextSegment | undefined,
	elementStyle: TextStyle | undefined,
	fallbackColor = '#000000',
): WarpRunPaint {
	const s: TextStyle = segment?.style ?? {};
	const family = s.fontFamily || elementStyle?.fontFamily;
	const decorations = [
		s.underline || s.hyperlink ? 'underline' : '',
		s.strikethrough ? 'line-through' : '',
	].filter(Boolean);
	return {
		fill: normalizeHexColor(
			s.color || elementStyle?.color,
			s.hyperlink ? HYPERLINK_COLOR : fallbackColor,
		),
		fontFamily: family ? getSubstituteFontFamily(family) : DEFAULT_FONT_FAMILY,
		fontSizePx: s.fontSize ?? elementStyle?.fontSize ?? 18,
		fontWeight: fontWeightAttribute(family, s.bold ?? elementStyle?.bold),
		fontStyle: resolveFontItalic(family, s.italic ?? elementStyle?.italic) ? 'italic' : undefined,
		textDecoration: decorations.length > 0 ? decorations.join(' ') : undefined,
	};
}
