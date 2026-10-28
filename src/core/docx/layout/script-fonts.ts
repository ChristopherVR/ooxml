import { segmentByFont } from '../../text/script-fonts';
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
	const pieces = segmentByFont(run.text, fonts);
	if (pieces.length === 1 && pieces[0]?.fontFamily === latin) return [run];
	const { breakAfter, ...style } = run;
	return pieces.map((piece, index) => ({
		...style,
		text: piece.text,
		fontFamily: piece.fontFamily,
		sourceLength: piece.text.length,
		...(breakAfter && index === pieces.length - 1 ? { breakAfter } : {}),
	}));
}
