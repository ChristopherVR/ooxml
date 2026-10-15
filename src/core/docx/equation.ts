import { XMLSerializer } from '@xmldom/xmldom';
import type { TextRun } from './model';
import { parseXml, type XmlDocument, type XmlElement } from './xml';

const MATH_NAMESPACES = new Set([
	'http://schemas.openxmlformats.org/officeDocument/2006/math',
	'http://purl.oclc.org/ooxml/officeDocument/math',
]);
const MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
type SerializableNode = Parameters<InstanceType<typeof XMLSerializer>['serializeToString']>[0];

export function isEquationElement(element: XmlElement): boolean {
	return (
		MATH_NAMESPACES.has(element.namespaceURI ?? '') &&
		(element.localName === 'oMath' || element.localName === 'oMathPara')
	);
}

/** Serializing the subtree supplies inherited namespaces, so consumers can parse it independently. */
export function parseEquation(element: XmlElement): TextRun | undefined {
	if (!isEquationElement(element)) return undefined;
	return {
		text: '',
		equation: {
			omml: new XMLSerializer().serializeToString(element as unknown as SerializableNode),
			display: element.localName === 'oMathPara',
		},
	};
}

/**
 * The equation's XML. An unchanged imported equation keeps its source subtree (extension
 * attributes included); an edited or new one is built from the run's OMML, wrapped in
 * `m:oMathPara` when `display` and unwrapped to its `m:oMath` when not.
 */
export function writeEquation(
	doc: XmlDocument,
	run: TextRun,
	base: TextRun | undefined,
	old: XmlElement | undefined,
): XmlElement {
	const equation = run.equation;
	if (!equation || run.text !== '') throw new Error('An equation run holds no text.');
	if (
		old &&
		isEquationElement(old) &&
		base?.equation?.omml === equation.omml &&
		base.equation.display === equation.display
	)
		return old.cloneNode(true) as XmlElement;
	let source: XmlElement;
	try {
		source = parseXml(equation.omml).documentElement;
	} catch (cause) {
		throw new Error(
			`Equation markup could not be read: ${cause instanceof Error ? cause.message : String(cause)}`,
			{ cause },
		);
	}
	if (!isEquationElement(source))
		throw new Error('Equation markup must be an m:oMath or m:oMathPara element.');
	let element = doc.importNode(source, true) as XmlElement;
	if (equation.display && element.localName === 'oMath') {
		const para = doc.createElementNS(MATH_NS, 'm:oMathPara');
		para.setAttributeNS('http://www.w3.org/2000/xmlns/', 'xmlns:m', MATH_NS);
		para.appendChild(element);
		element = para;
	} else if (!equation.display && element.localName === 'oMathPara') {
		const inner = Array.from(element.childNodes).find(
			(node): node is XmlElement =>
				node.nodeType === 1 && (node as XmlElement).localName === 'oMath',
		);
		if (inner) element = inner;
	}
	return element;
}
