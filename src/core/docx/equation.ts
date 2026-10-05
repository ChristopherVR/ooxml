import { XMLSerializer } from '@xmldom/xmldom';
import type { TextRun } from './model.js';
import type { XmlElement } from './xml.js';

const MATH_NAMESPACES = new Set([
	'http://schemas.openxmlformats.org/officeDocument/2006/math',
	'http://purl.oclc.org/ooxml/officeDocument/math',
]);
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

/** Equation source may be relocated within its paragraph, but cannot be created or changed. */
export function preserveEquation(
	run: TextRun,
	base: TextRun | undefined,
	old: XmlElement | undefined,
): XmlElement {
	if (
		!old ||
		!isEquationElement(old) ||
		!base?.equation ||
		run.text !== '' ||
		run.equation?.omml !== base.equation.omml ||
		run.equation?.display !== base.equation.display
	)
		throw new Error(
			'Equation editing or insertion is unsupported; imported OMML must remain unchanged.',
		);
	return old.cloneNode(true) as XmlElement;
}
