import { buildXml, isElement, parseXml, type XmlElement } from '../xml/index';
import { setSvgAttributes, SVG_NAMESPACE } from './svg-elements';

const ELEMENTS = new Set([
	'circle',
	'clipPath',
	'defs',
	'ellipse',
	'feColorMatrix',
	'feComponentTransfer',
	'feComposite',
	'feDropShadow',
	'feFlood',
	'feFuncB',
	'feFuncG',
	'feFuncR',
	'feGaussianBlur',
	'feMerge',
	'feMergeNode',
	'feMorphology',
	'filter',
	'g',
	'line',
	'linearGradient',
	'mask',
	'path',
	'polygon',
	'radialGradient',
	'rect',
	'stop',
	'text',
]);

// These presentation attributes cannot load resources. Paints and filter styles have separate
// value restrictions below because SVG and CSS also accept URLs in those positions.
const ATTRIBUTES = new Set([
	'color-interpolation-filters',
	'cx',
	'cy',
	'd',
	'dominant-baseline',
	'dx',
	'dy',
	'fill-opacity',
	'flood-opacity',
	'font-family',
	'font-size',
	'font-weight',
	'height',
	'id',
	'in',
	'in2',
	'intercept',
	'offset',
	'opacity',
	'operator',
	'points',
	'r',
	'radius',
	'result',
	'rx',
	'ry',
	'slope',
	'stdDeviation',
	'stop-opacity',
	'stroke-dasharray',
	'stroke-opacity',
	'stroke-width',
	'tableValues',
	'text-anchor',
	'transform',
	'type',
	'values',
	'width',
	'x',
	'x1',
	'x2',
	'y',
	'y1',
	'y2',
]);
const PAINT_ATTRIBUTES = new Set(['fill', 'stroke', 'flood-color', 'stop-color']);
const REFERENCE_ATTRIBUTES = new Set(['clip-path', 'filter', 'mask']);
const LOCAL_REFERENCE = /^url\(#[\w-]+\)$/u;
const COLOR = /^(?:#[\da-f]{3,8}|rgba?\([\d\s.,%]+\)|[a-z]+)$/iu;
const FILTER_STYLE =
	/^filter:(?:(?:blur|brightness|contrast|grayscale|hue-rotate|invert|opacity|saturate|sepia)\([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:%|px|deg)?\)\s*)+$/u;

export function safeSvgPreviewAttribute({ localName, namespaceURI, value }: Attr): boolean {
	if (namespaceURI) return false;
	if (PAINT_ATTRIBUTES.has(localName)) return COLOR.test(value) || LOCAL_REFERENCE.test(value);
	if (REFERENCE_ATTRIBUTES.has(localName)) return LOCAL_REFERENCE.test(value);
	if (localName === 'style') return FILTER_STYLE.test(value);
	return ATTRIBUTES.has(localName);
}

function safeChildren(parent: XmlElement): boolean {
	return Array.from(parent.childNodes).every((node) => {
		if (node.nodeType === 3) return true;
		return (
			isElement(node) &&
			node.namespaceURI === SVG_NAMESPACE &&
			!node.prefix &&
			ELEMENTS.has(node.localName) &&
			Array.from(node.attributes).every(safeSvgPreviewAttribute) &&
			safeChildren(node)
		);
	});
}

/** Render an inline gallery tile using only static SVG geometry and local paint definitions. */
export function renderSvgPreview(
	width: number,
	height: number,
	defs: string,
	body: string,
	viewBox: { width: number; height: number } = { width, height },
): string {
	if (![width, height, viewBox.width, viewBox.height].every(Number.isFinite)) return '';
	let svg: XmlElement;
	try {
		svg = parseXml(
			`<svg xmlns="${SVG_NAMESPACE}">${defs ? `<defs>${defs}</defs>` : ''}${body}</svg>`,
			{ warningsAsErrors: true },
		).documentElement;
	} catch {
		return '';
	}
	if (svg.attributes.length !== 1 || !safeChildren(svg)) return '';
	setSvgAttributes(svg, {
		width,
		height,
		viewBox: `0 0 ${viewBox.width} ${viewBox.height}`,
		'aria-hidden': 'true',
		focusable: 'false',
	});
	return buildXml(svg);
}
