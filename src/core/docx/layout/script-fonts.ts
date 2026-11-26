import { resolveFontForScript, segmentByScript } from '../../text/unicode-script-detection';
import type { RunFormatting } from '../run-style-model';
import type { ThemeCatalog } from '../theme-model';
import type { LayoutRun } from './input';
import { DEFAULT_FONT_FAMILY } from './units';

/** Split only literal source text; generated fields and objects keep their source mapping. */
export function scriptLayoutRuns(
	run: LayoutRun,
	formatting: RunFormatting,
	theme: ThemeCatalog | undefined,
): LayoutRun[] {
	if (
		!run.text ||
		run.hidden ||
		run.object ||
		run.footnoteReferenceId ||
		(run.sourceLength !== undefined && run.sourceLength !== run.text.length)
	)
		return [run];
	const latin = run.fontFamily || DEFAULT_FONT_FAMILY;
	const eaRole = formatting.fontTheme?.eastAsia;
	const csRole = formatting.fontTheme?.cs;
	const fonts = {
		latin,
		eastAsia: formatting.fontFamilyEastAsia || (eaRole && theme?.fonts[eaRole]?.eastAsia) || latin,
		complexScript:
			formatting.fontFamilyComplexScript ||
			(csRole && theme?.fonts[csRole]?.complexScript) ||
			latin,
		symbol: latin,
	};
	const complexSize = formatting.fontSizeComplexScript ?? run.fontSizePt;
	if (fonts.eastAsia === latin && fonts.complexScript === latin && complexSize === run.fontSizePt)
		return [run];
	const { breakAfter, ...style } = run;
	const pieces: LayoutRun[] = [];
	for (const piece of segmentByScript(run.text)) {
		const fontFamily = resolveFontForScript(piece.script, fonts) ?? latin;
		const fontSizePt = piece.script === 'complexScript' ? complexSize : run.fontSizePt;
		const previous = pieces.at(-1);
		if (previous && previous.fontFamily === fontFamily && previous.fontSizePt === fontSizePt) {
			previous.text += piece.text;
			previous.sourceLength = previous.text.length;
		} else {
			pieces.push({
				...style,
				text: piece.text,
				fontFamily,
				sourceLength: piece.text.length,
				...(fontSizePt !== undefined && { fontSizePt }),
			});
		}
	}
	const last = pieces.at(-1);
	if (pieces.length === 1 && last?.fontFamily === latin && last.fontSizePt === run.fontSizePt)
		return [run];
	if (last && breakAfter) last.breakAfter = breakAfter;
	return pieces;
}
