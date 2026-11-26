import {
	convertOmmlToLatex,
	convertOmmlToMathMl,
	EQUATION_PLACEHOLDER,
	fillPlaceholders,
	parseOmml,
} from 'ooxml-core/math';

export interface EquationPreview {
	latex: string;
	markup: string;
}

const PLACEHOLDER_ONLY = { 'm:oMath': { 'm:r': { 'm:t': EQUATION_PLACEHOLDER } } };
const EMPTY =
	/^(?:<m:oMathPara\b[^>]*>)?<m:oMath\b[^>]*(?:\/>|>\s*<\/m:oMath>)(?:<\/m:oMathPara>)?$/;

/**
 * MathML (and LaTeX) for an equation's markup. With `placeholders`, empty arguments and an empty
 * equation show the dotted box an editor shows while they wait for input.
 */
export function equationPreview(omml: string, placeholders = false): EquationPreview | undefined {
	try {
		let tree = parseOmml(omml);
		// A new equation holds nothing yet: it shows one placeholder until something is typed.
		if (placeholders) tree = EMPTY.test(omml) ? PLACEHOLDER_ONLY : fillPlaceholders(tree);
		const latex = convertOmmlToLatex(tree);
		const markup = convertOmmlToMathMl(tree);
		return { latex, markup };
	} catch {
		return undefined;
	}
}

/** MathML comes only from the fixed-tag converter, which escapes equation text and formats. */
export function renderEquation(
	doc: Document,
	preview: EquationPreview,
	display: boolean,
	label: string,
): Element | undefined {
	const parser = new (doc.defaultView?.DOMParser ?? DOMParser)();
	const math = parser.parseFromString(preview.markup, 'application/xml').documentElement;
	if (
		math.localName !== 'math' ||
		math.namespaceURI !== 'http://www.w3.org/1998/Math/MathML' ||
		!math.textContent?.trim()
	)
		return;
	math.setAttribute('display', display ? 'block' : 'inline');
	math.setAttribute('aria-label', preview.latex || label);
	return doc.importNode(math, true);
}
