/**
 * Professional and Linear: the two displays of an Office Math equation and the two input formats
 * (Unicode, LaTeX) its linear text is written in. Building up turns linear text into Office Math
 * structures; linearizing is the reverse. An equation shown in Linear is saved as its linear text
 * in one math run, which is how Word saves it too.
 */
import { convertLatexToOmml } from './latex-to-omml';
import { EQUATION_PLACEHOLDER, fillPlaceholders, stripPlaceholders } from './equation-placeholder';
import { MAX_LATEX_LENGTH } from './limits';
import { parseOmml } from './omml-dom';
import type { OmmlNode } from './omml-node';
import { convertOmmlToLatex } from './omml-to-latex';
import { contentKeys, ensureArr, readRunText } from './omml-to-latex-helpers';
import { convertOmmlToUnicodeMath } from './omml-to-unicode-math';
import { ommlToXml } from './omml-to-xml';
import { applyControlWords, unicodeMathToLatex } from './unicode-math';
import { MATH_AUTOCORRECT } from './unicode-math-symbols';

/** The notation linear text is written in: UnicodeMath (Word's default) or LaTeX. */
export type EquationInputFormat = 'unicode' | 'latex';

/** Markup holding an Office Math structure, not only runs. */
const STRUCTURE =
	/<m:(f|rad|nary|d|sSup|sSub|sSubSup|sPre|func|acc|m|eqArr|limLow|limUpp|bar|borderBox|groupChr)[\s/>]/;

function rootMath(tree: OmmlNode): OmmlNode | undefined {
	const para = tree['m:oMathPara'] as OmmlNode | undefined;
	const math = para ? ensureArr(para['m:oMath'])[0] : ensureArr(tree['m:oMath'])[0];
	return math;
}

function wrap(math: OmmlNode, display: boolean): string {
	return ommlToXml(display ? { 'm:oMathPara': { 'm:oMath': math } } : { 'm:oMath': math });
}

/** The markup of an equation that holds nothing yet. */
export function emptyEquationXml(display: boolean): string {
	return wrap({}, display);
}

/**
 * Office Math for linear text, built up into structures; placeholders still waiting for input
 * stay empty arguments. Undefined when LaTeX input cannot be converted (Unicode input always
 * builds: what it cannot structure stays text).
 */
export function buildUpEquation(
	text: string,
	format: EquationInputFormat,
	display: boolean,
): string | undefined {
	if (text.length > MAX_LATEX_LENGTH) return undefined;
	if (!text.trim()) return emptyEquationXml(display);
	try {
		const latex = format === 'unicode' ? unicodeMathToLatex(text) : text;
		const math = rootMath(stripPlaceholders(convertLatexToOmml(latex)));
		return math ? wrap(math, display) : emptyEquationXml(display);
	} catch {
		return undefined;
	}
}

/** Linear text for an equation's markup, with `⬚` in each empty argument. */
export function linearizeEquation(omml: string, format: EquationInputFormat): string {
	try {
		const tree = fillPlaceholders(parseOmml(omml));
		return format === 'unicode' ? convertOmmlToUnicodeMath(tree) : convertOmmlToLatex(tree);
	} catch {
		return '';
	}
}

/** Markup that shows `text` itself, unbuilt, as Word's Linear display saves it. */
export function linearEquationXml(text: string, display: boolean): string {
	const value = text.split(EQUATION_PLACEHOLDER).join('');
	return wrap(value ? { 'm:r': { 'm:t': value } } : {}, display);
}

/**
 * The equation's text when it is shown in Linear: its math holds only plain runs and building
 * that text up gives structures. Undefined for an equation shown in Professional.
 */
export function linearDisplayText(omml: string, format: EquationInputFormat): string | undefined {
	try {
		const math = rootMath(parseOmml(omml));
		if (!math) return undefined;
		const keys = contentKeys(math);
		if (!keys.length || keys.some((key) => !key.startsWith('m:r'))) return undefined;
		const text = keys
			.flatMap((key) => ensureArr(math[key]))
			.map(readRunText)
			.join('');
		const built = buildUpEquation(text, format, false);
		return built && STRUCTURE.test(built) ? text : undefined;
	} catch {
		return undefined;
	}
}

/**
 * Math AutoCorrect as the space after a control word triggers it: when the text before `caret`
 * ends in a known `\word`, the word becomes its character. Returns the new text and caret, or
 * undefined when nothing before the caret is a control word.
 */
export function mathAutoCorrect(
	text: string,
	caret: number,
): { text: string; caret: number } | undefined {
	const match = /\\([A-Za-z]+)$/.exec(text.slice(0, caret));
	const char = match && MATH_AUTOCORRECT[match[1]!];
	if (!match || !char) return undefined;
	const start = caret - match[0].length;
	return { text: text.slice(0, start) + char + text.slice(caret), caret: start + char.length };
}

export { applyControlWords };
