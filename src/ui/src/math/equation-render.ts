import { convertOmmlToLatex, convertOmmlToMathMl, parseOmml } from 'ooxml-core/math';

export interface EquationPreview {
	latex: string;
	markup: string;
}

export function equationPreview(omml: string): EquationPreview | undefined {
	try {
		const tree = parseOmml(omml);
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
