/** Per-script span decoration adapter. Font decisions live in core/text. */
import { segmentByFont, type ScriptFontSet } from 'ooxml-core/text';
import type { RunStyle } from './text-run-style';
export { resolveScriptFontSet, type ScriptFontSet, type ScriptFontFields } from 'ooxml-core/text';

export interface ScriptFontPiece {
	text: string;
	style?: RunStyle;
}

/** Repeat non-inherited text decorations on each span that changes its font. */
export function splitRunByScriptFont(
	text: string,
	fonts: ScriptFontSet,
	baseFontFamily: string,
	nestedDecoration?: RunStyle,
): ScriptFontPiece[] | undefined {
	const pieces = segmentByFont(text, fonts);
	if (!pieces.length || pieces.every((piece) => piece.fontFamily === baseFontFamily))
		return undefined;
	return pieces.map((piece) =>
		piece.fontFamily === baseFontFamily
			? { text: piece.text }
			: { text: piece.text, style: { ...nestedDecoration, fontFamily: piece.fontFamily } },
	);
}
