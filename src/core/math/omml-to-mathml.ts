/** Pure OMML to MathML conversion, with no browser or document-format dependency. */
import type { OmmlNode } from './omml-node.js';
import { stripXmlOrderSuffix } from './omml-node.js';
import { getOmmlMathColor, getOmmlMathFontSize } from './omml-color.js';
import { ensureArray } from './omml-mathml-helpers.js';
export type { OmmlNode } from './omml-node.js';
import {
	convertRun,
	convertFraction,
	convertRadical,
	convertSuperscript,
	convertSubscript,
	convertSubSup,
	convertPreSubSup,
	convertNary,
	convertDelimiter,
	convertMatrix,
	convertAccent,
	convertBar,
	convertLimLow,
	convertLimUpp,
	convertGroupChr,
	convertEqArr,
	convertBox,
	convertFunc,
} from './omml-mathml-converters.js';
// ── Dispatch ────────────────────────────────────────────────────────────────

/** Convert all child elements of an OMML container to MathML. */
function convertChildren(node: OmmlNode): string {
	if (!node || typeof node !== 'object') {
		return '';
	}
	const parts: string[] = [];

	for (const key of Object.keys(node)) {
		if (key.startsWith('@_')) {
			continue;
		}
		// Keys may carry `#pptx-order-N` position markers (interleaved sibling
		// sequences); strip them before dispatching on the tag name.
		const tag = stripXmlOrderSuffix(key);
		if (tag === 'm:oMathPara') {
			continue;
		}

		const items = ensureArray(node[key]);
		for (const item of items) {
			const result = convertElement(tag, item);
			if (result) {
				parts.push(result);
			}
		}
	}

	return parts.join('');
}

/** Convert a single OMML element by tag name. */
function convertElement(tag: string, node: OmmlNode): string {
	switch (tag) {
		case 'm:r':
			return convertRun(node);
		case 'm:f':
			return convertFraction(node, convertChildren);
		case 'm:rad':
			return convertRadical(node, convertChildren);
		case 'm:sSup':
			return convertSuperscript(node, convertChildren);
		case 'm:sSub':
			return convertSubscript(node, convertChildren);
		case 'm:sSubSup':
			return convertSubSup(node, convertChildren);
		case 'm:sPre':
			return convertPreSubSup(node, convertChildren);
		case 'm:nary':
			return convertNary(node, convertChildren);
		case 'm:d':
			return convertDelimiter(node, convertChildren);
		case 'm:m':
			return convertMatrix(node, convertChildren);
		case 'm:acc':
			return convertAccent(node, convertChildren);
		case 'm:bar':
			return convertBar(node, convertChildren);
		case 'm:limLow':
			return convertLimLow(node, convertChildren);
		case 'm:limUpp':
			return convertLimUpp(node, convertChildren);
		case 'm:groupChr':
			return convertGroupChr(node, convertChildren);
		case 'm:eqArr':
			return convertEqArr(node, convertChildren);
		case 'm:box':
			return convertBox(node, convertChildren);
		case 'm:borderBox':
			return convertBox(node, convertChildren);
		case 'm:func':
			return convertFunc(node, convertChildren);
		case 'm:oMath':
			return `<mrow>${convertChildren(node)}</mrow>`;
		default:
			// TODO: defer exotic constructs (m:phant, m:sSubSupPr edge cases,
			// m:argPr scaling, etc.) : passthrough as empty for now.
			return '';
	}
}

/** Locate all `m:oMath` root elements inside an OMML wrapper node. */
function findOmathRoots(node: OmmlNode): OmmlNode[] {
	if (node['m:oMath']) {
		return ensureArray(node['m:oMath']);
	}
	const para = node['m:oMathPara'];
	if (para) {
		const paraNode = Array.isArray(para) ? (para[0] as OmmlNode) : (para as OmmlNode);
		if (paraNode['m:oMath']) {
			return ensureArray(paraNode['m:oMath']);
		}
	}
	const contentTags = new Set(['m:r', 'm:f', 'm:rad', 'm:sSup', 'm:sSub', 'm:box']);
	if (Object.keys(node).some((key) => contentTags.has(stripXmlOrderSuffix(key)))) {
		return [node];
	}
	return [];
}

// ── Public API ──────────────────────────────────────────────────────────────

/**
 * Convert an OMML XML node (from fast-xml-parser) into a MathML string.
 *
 * Accepts the object at the `<a14:m>` / `<m:oMathPara>` level or directly at
 * `<m:oMath>`. Returns a `<math>` element string, or empty string if the input
 * is empty / unparseable.
 */
export function convertOmmlToMathMl(ommlNode: OmmlNode): string {
	if (!ommlNode || typeof ommlNode !== 'object') {
		return '';
	}

	const oMaths = findOmathRoots(ommlNode);
	if (oMaths.length === 0) {
		return '';
	}

	const innerParts = oMaths.map((om) => convertChildren(om));
	const inner = innerParts.join('');
	if (inner.length === 0) {
		return '';
	}

	const color = getOmmlMathColor(ommlNode);
	const colorAttribute = color ? ` mathcolor="${color}"` : '';
	const fontSize = getOmmlMathFontSize(ommlNode);
	const sizeAttribute = fontSize ? ` mathsize="${fontSize}pt"` : '';
	return `<math xmlns="http://www.w3.org/1998/Math/MathML" display="inline"${colorAttribute}${sizeAttribute}>${inner}</math>`;
}

/**
 * Convenience alias used by the Vue viewer: `ommlToMathml(omml)`.
 *
 * Accepts the parsed OMML object (the shape stored on
 * {@link TextSegment.equationXml}) or a raw OMML markup string. String inputs
 * are not re-parsed here (no XML parser dependency in this pure module) : they
 * are returned wrapped so callers can still surface raw markup if needed.
 */
export function ommlToMathml(omml: OmmlNode | string): string {
	if (typeof omml === 'string') {
		// A bare string cannot be meaningfully walked without a parser; the
		// viewer always passes the parsed object. Return empty to stay pure.
		return '';
	}
	return convertOmmlToMathMl(omml);
}
