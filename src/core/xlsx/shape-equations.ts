import { checkEquationTree, MAX_OMML_LENGTH } from '../math/limits';
import { NS, children, elements, first, parseXml, type XmlElement } from '../xml/index';
import { outerXml } from './read/xml-util';

const DRAWING_MATH = 'http://schemas.microsoft.com/office/drawing/2010/main';
const STRICT_MATH = 'http://purl.oclc.org/ooxml/officeDocument/math';

export interface ShapeEquation {
	/** Standalone Office Math from the a14:m wrapper. Missing when over the conversion limit. */
	omml?: string;
	fallback: string;
}

type ShapeTextItem = { kind: 'text'; text: string } | ({ kind: 'equation' } & ShapeEquation);
export type EquationParagraphs = ShapeTextItem[][];

const mathWrapper = (node: XmlElement) =>
	node.localName === 'm' && node.namespaceURI === DRAWING_MATH;
const isMath = (node: XmlElement) =>
	(node.namespaceURI === NS.m || node.namespaceURI === STRICT_MATH) &&
	(node.localName === 'oMath' || node.localName === 'oMathPara');

const paragraphs = (shape: XmlElement | undefined) => {
	const body = first(shape, 'txBody', NS.xdr);
	return body ? children(body, 'p', NS.a) : [];
};

const text = (node: XmlElement) =>
	node.localName === 'br' ? '\n' : (first(node, 't', NS.a)?.textContent ?? '');

/** Excel stores math in the Choice's text box and a linear drawing run in its Fallback. */
export function equationShapeBranches(element: XmlElement) {
	if (element.namespaceURI !== NS.mc || element.localName !== 'AlternateContent') return;
	for (const choice of children(element, 'Choice', NS.mc)) {
		const requires = choice.getAttribute('Requires')?.trim().split(/\s+/) ?? [];
		if (requires.length !== 1 || choice.lookupNamespaceURI(requires[0] ?? '') !== DRAWING_MATH)
			continue;
		const shape = first(choice, 'sp', NS.xdr);
		if (shape && paragraphs(shape).some((paragraph) => elements(paragraph).some(mathWrapper)))
			return { shape, fallback: first(first(element, 'Fallback', NS.mc), 'sp', NS.xdr) };
	}
}

function equationXml(wrapper: XmlElement): string | undefined {
	const math = elements(wrapper).find(isMath);
	if (!math) return;
	try {
		checkEquationTree(math);
		const xml = outerXml(math);
		return xml.length <= MAX_OMML_LENGTH ? xml : undefined;
	} catch {
		return undefined;
	}
}

const textRun = (node: XmlElement) =>
	node.namespaceURI === NS.a && ['r', 'br', 'fld'].includes(node.localName);

export function readEquationParagraphs(
	shape: XmlElement,
	fallback?: XmlElement,
): EquationParagraphs | undefined {
	const alternate = paragraphs(fallback);
	let hasEquation = false;
	const result = paragraphs(shape).map((paragraph, index) => {
		const runs = elements(paragraph).filter((node) => mathWrapper(node) || textRun(node));
		const other = alternate[index];
		const fallbackRuns = other ? elements(other).filter(textRun) : [];
		return runs.map((node, runIndex): ShapeTextItem => {
			if (!mathWrapper(node)) return { kind: 'text', text: text(node) };
			hasEquation = true;
			const omml = equationXml(node);
			const fallback = fallbackRuns.length === runs.length ? fallbackRuns[runIndex] : undefined;
			return {
				kind: 'equation',
				...(omml ? { omml } : {}),
				fallback: fallback ? text(fallback) : '',
			};
		});
	});
	return hasEquation ? result : undefined;
}

/** Edits replace one OMML element and its matching fallback run, preserving adjacent content. */
export function replaceShapeEquation(
	anchorXml: string,
	index: number,
	value: Required<ShapeEquation>,
) {
	const doc = parseXml(anchorXml);
	for (const content of elements(doc.documentElement)) {
		const branches = equationShapeBranches(content);
		const choice =
			branches?.shape ??
			(content.namespaceURI === NS.xdr && content.localName === 'sp' ? content : undefined);
		if (!choice) continue;
		const fallback = paragraphs(branches?.fallback);
		for (const [p, paragraph] of paragraphs(choice).entries()) {
			const runs = elements(paragraph).filter((node) => mathWrapper(node) || textRun(node));
			for (const [position, run] of runs.entries()) {
				if (!mathWrapper(run) || index-- !== 0) continue;
				const old = elements(run).find(isMath);
				if (!old || !equationXml(run)) return;
				const math = parseShapeEquation(value);
				run.replaceChild(doc.importNode(math, true), old);
				const other = fallback[p];
				const fallbacks = other ? elements(other).filter(textRun) : [];
				const match =
					fallbacks.length === runs.length ? first(fallbacks[position], 't', NS.a) : undefined;
				if (match) match.textContent = value.fallback;
				return { doc, choice, fallback: branches?.fallback };
			}
		}
	}
}

export function parseShapeEquation(value: Required<ShapeEquation>): XmlElement {
	if (value.omml.length > MAX_OMML_LENGTH || value.fallback.length > MAX_OMML_LENGTH)
		throw new RangeError('OMML exceeds the size limit');
	const root = parseXml(value.omml, { label: 'XLSX equation' }).documentElement;
	if (!isMath(root)) throw new Error('Expected an Office Math equation');
	checkEquationTree(root);
	return root;
}

/** New text boxes use the same two shape branches as Excel's Insert Equation. */
export function wrapEquationShape(shape: string, value: Required<ShapeEquation>): string {
	const body =
		'<xdr:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p>';
	const end = '<a:endParaRPr sz="1100"/></a:p></xdr:txBody></xdr:sp>';
	const math = outerXml(parseShapeEquation(value));
	const fallback = value.fallback
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;');
	return `<mc:AlternateContent xmlns:mc="${NS.mc}"><mc:Choice xmlns:a14="${DRAWING_MATH}" Requires="a14">${shape}${body}<a14:m>${math}</a14:m>${end}</mc:Choice><mc:Fallback>${shape}${body}<a:r><a:rPr sz="1100"><a:latin typeface="Cambria Math"/></a:rPr><a:t>${fallback}</a:t></a:r>${end}</mc:Fallback></mc:AlternateContent>`;
}
