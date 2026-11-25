import { RasterImageError, checkImageDimensions, type ImageLimits } from './raster-image';
import { SVG_NAMESPACE } from './svg-elements';
import { safeSvgPreviewAttribute } from './svg-preview';
import { buildXml, isElement, parseXml, type XmlElement, XmlLimitError } from '../xml/index';

const MAX_SVG_BYTES = 256 * 1024;
const ELEMENTS = new Set([
	'svg',
	'g',
	'defs',
	'path',
	'rect',
	'circle',
	'ellipse',
	'line',
	'polyline',
	'polygon',
	'linearGradient',
	'radialGradient',
	'stop',
	'text',
	'tspan',
	'title',
	'desc',
]);
const STATIC_ATTRIBUTES = new Set([
	'version',
	'viewBox',
	'preserveAspectRatio',
	'gradientUnits',
	'gradientTransform',
	'spreadMethod',
	'fill-rule',
	'stroke-linecap',
	'stroke-linejoin',
	'stroke-miterlimit',
	'stroke-dashoffset',
	'vector-effect',
	'font-style',
	'text-decoration',
	'color',
]);
const LOCAL_PAINT = /^url\(#([A-Za-z_][\w.-]{0,127})\)$/;
const XML_DECLARATION =
	/^<\?xml\s+version\s*=\s*(?:"1\.0"|'1\.0')(?:\s+encoding\s*=\s*(?:"UTF-8"|'UTF-8'))?(?:\s+standalone\s*=\s*(?:"(?:yes|no)"|'(?:yes|no)'))?\s*\?>/i;
const WIDTH_UNIT: Readonly<Record<string, number>> = {
	px: 1,
	in: 96,
	cm: 96 / 2.54,
	mm: 96 / 25.4,
	pt: 96 / 72,
	pc: 16,
};

function size(value: string | null): number | undefined {
	const match = /^([+]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?)(px|in|cm|mm|pt|pc)?$/i.exec(
		value?.trim() ?? '',
	);
	if (match) return Number(match[1]) * WIDTH_UNIT[match[2]?.toLowerCase() ?? 'px']!;
	return undefined;
}

function unsafeSvg(): never {
	throw new RasterImageError(
		'unsupported-image',
		'SVG picture contains active or unsupported content.',
	);
}

function checkElements(root: XmlElement): void {
	const gradients = new Set<string>();
	const ids = new Set<string>();
	const paints: string[] = [];
	const nodes = [root];
	for (const node of nodes) {
		if (node.namespaceURI !== SVG_NAMESPACE || node.prefix || !ELEMENTS.has(node.localName))
			unsafeSvg();
		if (node !== root && node.localName === 'svg') unsafeSvg();
		const id = node.getAttribute('id');
		if (id !== null) {
			if (ids.has(id)) unsafeSvg();
			ids.add(id);
			if (node.localName === 'linearGradient' || node.localName === 'radialGradient')
				gradients.add(id);
		}
		for (const attr of Array.from(node.attributes)) {
			if (node === root && attr.name === 'xmlns' && attr.value === SVG_NAMESPACE) continue;
			if (attr.namespaceURI) unsafeSvg();
			if ((attr.name === 'fill' || attr.name === 'stroke') && attr.value.startsWith('url')) {
				const paint = LOCAL_PAINT.exec(attr.value);
				if (!paint?.[1]) unsafeSvg();
				paints.push(paint[1]);
			} else if (
				['style', 'clip-path', 'mask', 'filter'].includes(attr.name) ||
				(!STATIC_ATTRIBUTES.has(attr.name) && !safeSvgPreviewAttribute(attr)) ||
				(/^(fill|stroke|stop-color|flood-color)$/.test(attr.name) && /url/i.test(attr.value))
			)
				unsafeSvg();
		}
		for (const child of Array.from(node.childNodes)) {
			if (isElement(child)) nodes.push(child);
			else if (![3, 4, 8].includes(child.nodeType)) unsafeSvg();
		}
	}
	if (paints.some((id) => !gradients.has(id))) unsafeSvg();
}

/** Verify a bounded static SVG and serialize its root without a prolog or processing instructions. */
export function inspectSvgPicture(bytes: Uint8Array, limits: ImageLimits) {
	if (bytes.length > MAX_SVG_BYTES)
		throw new RasterImageError('image-limit', 'SVG picture exceeds the 256 KiB limit.');
	let root: XmlElement;
	try {
		const source = new TextDecoder('utf-8', { fatal: true })
			.decode(bytes)
			.replace(XML_DECLARATION, '');
		const svg = parseXml(source, {
			label: 'SVG',
			warningsAsErrors: true,
			limits: { maxDepth: 32, maxNodes: 4096 },
		});
		if (Array.from(svg.childNodes).some((node) => node.nodeType === 7)) unsafeSvg();
		root = svg.documentElement;
	} catch (error) {
		if (error instanceof RasterImageError) throw error;
		if (error instanceof XmlLimitError)
			throw new RasterImageError('image-limit', 'SVG picture exceeds the markup complexity limit.');
		throw new RasterImageError('unsupported-image', 'Invalid SVG or unsupported picture format.');
	}
	if (root.localName !== 'svg' || root.namespaceURI !== SVG_NAMESPACE) unsafeSvg();
	checkElements(root);
	const rawBox = root.getAttribute('viewBox');
	const box = rawBox
		?.trim()
		.split(/[\s,]+/)
		.map(Number);
	if (
		box &&
		(box.length !== 4 ||
			box.some((value) => !Number.isFinite(value)) ||
			box[2]! <= 0 ||
			box[3]! <= 0)
	)
		throw new RasterImageError('invalid-image', 'SVG has an invalid viewBox.');
	const width = size(root.getAttribute('width'));
	const height = size(root.getAttribute('height'));
	const boxRatio = box ? box[2]! / box[3]! : 1;
	const pixelWidth = Math.ceil(
		width ?? (height === undefined ? (box?.[2] ?? 0) : height * boxRatio),
	);
	const pixelHeight = Math.ceil(
		height ?? (width === undefined ? (box?.[3] ?? 0) : width / boxRatio),
	);
	checkImageDimensions(pixelWidth, pixelHeight, limits);
	return { pixelWidth, pixelHeight, root };
}

export function svgPictureBytes(root: XmlElement): Uint8Array {
	const bytes = new TextEncoder().encode(buildXml(root));
	if (bytes.length > MAX_SVG_BYTES)
		throw new RasterImageError('image-limit', 'Normalized SVG picture exceeds the 256 KiB limit.');
	return bytes;
}
