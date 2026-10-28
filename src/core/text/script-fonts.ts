import { getSubstituteFontFamily, parsePanoseString } from './font-substitution';
import { resolveFontForScript, segmentByScript } from './unicode-script-detection';

/** Resolved family names or CSS stacks, chosen by the host's font policy. */
export interface ScriptFontSet {
	latin: string;
	eastAsia: string;
	complexScript: string;
	symbol: string;
}

/** DrawingML script declarations, independent of a product's text model. */
export interface ScriptFontFields {
	eastAsiaFont?: string | undefined;
	complexScriptFont?: string | undefined;
	symbolFont?: string | undefined;
	eastAsiaFontPanose?: string | undefined;
	complexScriptFontPanose?: string | undefined;
	symbolFontPanose?: string | undefined;
}

export function resolveScriptFontSet(
	run: ScriptFontFields | undefined,
	body: ScriptFontFields | undefined,
	baseFontFamily: string,
): ScriptFontSet {
	const resolve = (name: string | undefined, panose: string | undefined) =>
		name ? getSubstituteFontFamily(name, parsePanoseString(panose)) : baseFontFamily;
	return {
		latin: baseFontFamily,
		eastAsia: resolve(
			run?.eastAsiaFont || body?.eastAsiaFont,
			run?.eastAsiaFontPanose ?? body?.eastAsiaFontPanose,
		),
		complexScript: resolve(
			run?.complexScriptFont || body?.complexScriptFont,
			run?.complexScriptFontPanose ?? body?.complexScriptFontPanose,
		),
		symbol: resolve(
			run?.symbolFont || body?.symbolFont,
			run?.symbolFontPanose ?? body?.symbolFontPanose,
		),
	};
}

export interface ResolvedFontRun {
	text: string;
	/** UTF-16 source offset, including surrogate pairs. */
	startIndex: number;
	fontFamily: string;
}

/** Shared script selection; equal adjacent font choices remain one run. */
export function segmentByFont(text: string, fonts: ScriptFontSet): ResolvedFontRun[] {
	const result: ResolvedFontRun[] = [];
	for (const run of segmentByScript(text)) {
		const fontFamily = resolveFontForScript(run.script, fonts) ?? fonts.latin;
		const previous = result.at(-1);
		if (previous?.fontFamily === fontFamily) previous.text += run.text;
		else result.push({ text: run.text, startIndex: run.startIndex, fontFamily });
	}
	return result;
}
