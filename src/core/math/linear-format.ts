/**
 * Professional and Linear: the two displays of an Office Math equation and the two input formats
 * (Unicode, LaTeX) its linear text is written in. Building up turns linear text into Office Math
 * structures; linearizing is the reverse. An equation shown in Linear is saved as its linear
 * (UnicodeMath) text in math runs, each build-up operator in a run of its own, as Word saves it.
 *
 * Linear text cannot hold everything Office Math can (mixed run formatting, several equations in
 * one display paragraph, properties it has no spelling for), so `editableEquation` only offers an
 * equation for editing when its text rebuilds it exactly; anything else stays as it is.
 */
import {
	canonicalMath,
	coalesceRuns,
	controlFormat,
	onlyMathAndWord,
	orderedEntries,
	sharedRunFormat,
	withRunFormat,
} from './equation-canonical';
import {
	convertLatexToOmml,
	convertOmmlToLatex,
	ensureArr,
	fillPlaceholders,
	MAX_LATEX_LENGTH,
	mergeSiblings,
	ommlToXml,
	parseOmml,
	readRunText,
	type OmmlNode,
} from './index';
import { stripPlaceholders } from './equation-placeholder-strip';
import { convertOmmlToUnicodeMath } from './omml-to-unicode-math';
import { applyControlWords, unicodeMathToLatex } from './unicode-math';
import { isCombining, MATH_AUTOCORRECT } from './unicode-math-symbols';

/** The notation linear text is written in: UnicodeMath (Word's default) or LaTeX. */
export type EquationInputFormat = 'unicode' | 'latex';

/**
 * What an edited equation keeps from the original: its paragraph properties, its runs' format and
 * its control characters' format (the math font Word writes on both).
 */
export interface EquationFrame {
	paraProps?: OmmlNode;
	runFormat?: OmmlNode;
	controlFormat?: OmmlNode;
}

/** An equation the linear editor can edit without losing anything. */
export interface EditableEquation {
	/** Its linear text in the input format. */
	text: string;
	/** Whether it is shown in Linear rather than built up. */
	linear: boolean;
	/** Whether it can also be shown in Linear: its UnicodeMath text rebuilds it exactly. */
	canLinear: boolean;
	frame: EquationFrame;
}

/** Characters Word's Linear display saves in runs of their own. */
const OPERATOR_CHARS = new Set([
	...'/∕¦┬┴^_▒&@√∛∜■█□¯▁⏞⏟〖〗│()[]{}|‖⟨⟩⌊⌋⌈⌉├┤"\\⁡',
	...'∑∏∐∫∬∭∮⋃⋂⋁⋀',
]);
/** Accents (combining marks) and escapes are operators of the linear text too. */
const isOperator = (char: string) => OPERATOR_CHARS.has(char) || isCombining(char);

function wrap(math: OmmlNode, display: boolean, frame: EquationFrame = {}): string {
	const body = withRunFormat(math, frame.runFormat, frame.controlFormat);
	if (!display) return ommlToXml({ 'm:oMath': body });
	const para: OmmlNode = frame.paraProps ? { 'm:oMathParaPr': frame.paraProps } : {};
	para['m:oMath'] = body;
	return ommlToXml({ 'm:oMathPara': para });
}

/** The markup of an equation that holds nothing yet. */
export function emptyEquationXml(display: boolean): string {
	return wrap({}, display);
}

/** The `m:oMath` content linear text builds up to, runs coalesced and placeholders empty. */
function buildUpMath(text: string, format: EquationInputFormat): OmmlNode | undefined {
	if (text.length > MAX_LATEX_LENGTH) return undefined;
	if (!text.trim()) return {};
	try {
		const latex = format === 'unicode' ? unicodeMathToLatex(text) : text;
		const tree = stripPlaceholders(convertLatexToOmml(latex));
		const para = tree['m:oMathPara'] as OmmlNode | undefined;
		return coalesceRuns(ensureArr(para?.['m:oMath'])[0] ?? {});
	} catch {
		return undefined;
	}
}

/**
 * Office Math for linear text, built up into structures; placeholders still waiting for input
 * stay empty arguments. Undefined when LaTeX input cannot be converted or the text nests too
 * deeply (Unicode input otherwise always builds: what it cannot structure stays text).
 */
export function buildUpEquation(
	text: string,
	format: EquationInputFormat,
	display: boolean,
	frame?: EquationFrame,
): string | undefined {
	const math = buildUpMath(text, format);
	return math && wrap(math, display, frame);
}

/** The single `m:oMath` of an equation and its paragraph properties; undefined for several. */
function soleMath(tree: OmmlNode): { math: OmmlNode; paraProps?: OmmlNode } | undefined {
	const para = tree['m:oMathPara'];
	if (para === undefined) {
		const maths = ensureArr(tree['m:oMath']);
		return maths.length === 1 ? { math: maths[0]! } : undefined;
	}
	let math: OmmlNode | undefined;
	let paraProps: OmmlNode | undefined;
	for (const [tag, value] of orderedEntries(ensureArr(para)[0] ?? {})) {
		if (tag === 'm:oMathParaPr') paraProps = value as OmmlNode;
		else if (tag === 'm:oMath' && !math) math = value as OmmlNode;
		else return undefined;
	}
	return math ? { math, ...(paraProps ? { paraProps } : {}) } : undefined;
}

const linearize = (math: OmmlNode, format: EquationInputFormat) => {
	const tree = fillPlaceholders({ 'm:oMath': math });
	return format === 'unicode' ? convertOmmlToUnicodeMath(tree) : convertOmmlToLatex(tree);
};

/** Linear text for an equation's markup, with `⬚` in each empty argument; undefined when unread. */
export function linearizeEquation(omml: string, format: EquationInputFormat): string | undefined {
	try {
		const sole = soleMath(parseOmml(omml));
		return sole && linearize(sole.math, format);
	} catch {
		return undefined;
	}
}

/** Markup that shows UnicodeMath `text` itself, unbuilt, as Word's Linear display saves it. */
export function linearEquationXml(text: string, display: boolean, frame?: EquationFrame): string {
	const runs: string[] = [];
	// Placeholders stay in the text (Word shows them as ⬚), so empty arguments read back.
	for (const char of text) {
		const last = runs.length - 1;
		if (!isOperator(char) && last >= 0 && !isOperator(runs[last]!)) runs[last] += char;
		else runs.push(char);
	}
	return wrap(mergeSiblings(runs.map((t) => ({ 'm:r': { 'm:t': t } }))), display, frame);
}

const same = (a: OmmlNode | undefined, b: OmmlNode) =>
	a !== undefined && canonicalMath(a) === canonicalMath(b);

/** Whether UnicodeMath `text` rebuilds `math` exactly. */
const unicodeHolds = (text: string, math: OmmlNode) => same(buildUpMath(text, 'unicode'), math);

/**
 * Linear text of `math` when it is shown in Linear: plain runs with an operator run of its own,
 * whose text builds up into something else (text that rebuilds to the same runs is Professional).
 */
function linearText(math: OmmlNode): string | undefined {
	const entries = orderedEntries(math);
	if (!entries.length || entries.some(([tag]) => tag !== 'm:r')) return undefined;
	const runs = entries.map(([, run]) => run as OmmlNode);
	const plain = (run: OmmlNode) =>
		Object.keys(run).every((key) => key === 'm:t' || key === 'w:rPr');
	if (!runs.every(plain)) return undefined;
	const texts = runs.map(readRunText);
	if (!texts.some((text) => isOperator(text))) return undefined;
	const text = texts.join('');
	return unicodeHolds(text, math) ? undefined : text;
}

/**
 * The equation's UnicodeMath text when it is shown in Linear (as Word saves it: plain runs,
 * build-up operators in runs of their own); undefined for an equation shown in Professional.
 * The input format does not matter: Linear text is always stored as UnicodeMath.
 */
export function linearDisplayText(omml: string): string | undefined {
	try {
		const sole = soleMath(parseOmml(omml));
		return sole && linearText(sole.math);
	} catch {
		return undefined;
	}
}

/**
 * The equation as linear text the editor can change, when that text rebuilds it exactly (its
 * canonical form, see `canonicalMath`); undefined for an equation linear text would change:
 * several equations in one paragraph, mixed formatting, other namespaces, control characters with
 * more than the math font, or properties its notation cannot spell.
 */
export function editableEquation(
	omml: string,
	format: EquationInputFormat,
): EditableEquation | undefined {
	try {
		const tree = parseOmml(omml);
		const sole = onlyMathAndWord(tree) ? soleMath(tree) : undefined;
		if (!sole) return undefined;
		const runFormat = sharedRunFormat(sole.math);
		const control = controlFormat(sole.math);
		if (runFormat === 'mixed' || control === 'other') return undefined;
		const frame: EquationFrame = {
			...(sole.paraProps ? { paraProps: sole.paraProps } : {}),
			...(runFormat ? { runFormat } : {}),
			...(control ? { controlFormat: control } : {}),
		};
		const shown = linearText(sole.math);
		if (shown !== undefined) {
			if (format === 'unicode') return { text: shown, linear: true, canLinear: true, frame };
			// Linear text is kept as UnicodeMath; in LaTeX it is offered only when it reads back.
			const built = buildUpMath(shown, 'unicode');
			const text = built && linearize(built, 'latex');
			if (!built || !text || !same(buildUpMath(text, 'latex'), built)) return undefined;
			return { text, linear: true, canLinear: true, frame };
		}
		const text = linearize(sole.math, format);
		if (!same(buildUpMath(text, format), sole.math)) return undefined;
		const canLinear =
			format === 'unicode' || unicodeHolds(linearize(sole.math, 'unicode'), sole.math);
		return { text, linear: false, canLinear, frame };
	} catch {
		return undefined;
	}
}

/**
 * Markup for edited linear text: built up, or kept linear (as UnicodeMath) when `linear`.
 * LaTeX text shown in Linear is written as its UnicodeMath, only when that rebuilds it exactly;
 * undefined otherwise (or when the text cannot be read).
 */
export function writeEditedEquation(edit: {
	text: string;
	format: EquationInputFormat;
	display: boolean;
	linear: boolean;
	frame?: EquationFrame;
}): string | undefined {
	const { text, format, display, linear, frame } = edit;
	if (!linear) return buildUpEquation(text, format, display, frame);
	let unicode: string | undefined = text;
	if (format === 'latex') {
		const math = buildUpMath(text, 'latex');
		unicode = math && linearize(math, 'unicode');
		if (!math || unicode === undefined || !unicodeHolds(unicode, math)) return undefined;
	}
	const xml = linearEquationXml(unicode, display, frame);
	// Linear text that builds up must read back as Linear, or it would be saved as plain text.
	return linearDisplayText(xml) !== undefined || unicodeHolds(unicode, inner(xml))
		? xml
		: undefined;
}

const inner = (xml: string): OmmlNode => soleMath(parseOmml(xml))?.math ?? {};

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
