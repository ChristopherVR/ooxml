/**
 * Convert one static page SVG (from `exportPageSvg`) into a vector PDF page: paths with fill,
 * stroke, dash, caps and joins, arrowhead markers, rect clips, real text and embedded images.
 * What the PDF cannot draw degrades and is listed (`degraded`); what it cannot represent at all
 * throws `PdfFallback` inside `convertSvgPage`, which then reports `fallback` so the caller can
 * use a picture of the page instead.
 */
import type { PdfFontSet } from './pdf-embedded-font';
import { PdfVectorPage, type PdfImageResource } from './pdf-vector';
import { pathEnds, pathOperators, parsePath, type PathSegment } from './pdf-path';
import { drawSvgText, type SvgFontFamilyMapper } from './pdf-svg-text';
import { gradientPaint, isGradient, pathBounds, type GradientPaint } from './pdf-svg-gradient';
import { PDF_SVG_NOTES, type PdfSvgNoteCode } from './pdf-svg-notes';
import { childElements, descendantElements, elementId } from './pdf-svg-dom';
import { basicShapePath } from './pdf-svg-shapes';
import { paintSvgShadow, readShadowFilter } from './pdf-svg-shadow';
import {
	hrefOf,
	number,
	parseColor,
	parseTransform,
	PdfFallback,
	urlId,
	type Rgba,
} from './pdf-svg-style';

export interface SvgPageConversion {
	/** The vector page, or `undefined` when `fallback` lists why it cannot be vector. */
	page?: PdfVectorPage;
	fallback: string[];
	degraded: string[];
}

const MAX_NODES = 250_000;
const SKIPPED = new Set([
	'defs',
	'title',
	'desc',
	'metadata',
	'clipPath',
	'marker',
	'symbol',
	'linearGradient',
	'radialGradient',
	'stop',
	'filter',
	'mask',
	'style',
]);

/** Every distinct image `href` (data URL) the page draws, for the caller to decode. */
export function collectSvgImages(svg: Element): string[] {
	const found = new Set<string>();
	for (const node of descendantElements(svg)) {
		if (node.localName !== 'use' && node.localName !== 'image') continue;
		const image = imageHref(svg, node);
		if (image) found.add(image);
	}
	return [...found];
}

function elementById(root: Element, id: string): Element | undefined {
	for (const node of descendantElements(root))
		if (node.hasAttribute('id') && elementId(node) === id) return node;
	return undefined;
}

function imageHref(root: Element, node: Element): string | undefined {
	const href = hrefOf(node);
	if (!href) return undefined;
	if (!href.startsWith('#')) return href;
	const target = elementById(root, href.slice(1));
	const image =
		target?.localName === 'image'
			? target
			: target && descendantElements(target).find((child) => child.localName === 'image');
	return image ? (hrefOf(image) ?? undefined) : undefined;
}

class Converter {
	readonly degraded = new Set<PdfSvgNoteCode>();
	#nodes = 0;
	constructor(
		readonly root: Element,
		readonly page: PdfVectorPage,
		readonly images: ReadonlyMap<string, PdfImageResource>,
		readonly fonts?: PdfFontSet,
		readonly familyOf?: SvgFontFamilyMapper,
	) {}
	/** Set while a glow group is drawn: the halo width in the group's units. */
	#glow: number | undefined;
	#ids = new Map<string, Element>();
	#byId(id: string | undefined): Element | undefined {
		if (!id) return undefined;
		if (!this.#ids.size)
			for (const node of descendantElements(this.root))
				if (node.hasAttribute('id')) this.#ids.set(elementId(node), node);
		return this.#ids.get(id);
	}
	/** A flat colour, or (for fills with `bounds`) a gradient shading. */
	#paint(
		value: string | null,
		bounds?: [number, number, number, number],
		gradient = false,
	): Rgba | GradientPaint | null {
		const id = urlId(value);
		if (!id) return parseColor(value);
		const target = this.#byId(id);
		if (isGradient(target)) {
			if (gradient) {
				const paint = gradientPaint(target, bounds);
				if (paint) {
					if (paint.variableOpacity) this.degraded.add('gradientOpacityAveraged');
					if (paint.spreadApproximated) this.degraded.add('gradientSpreadApproximated');
					return paint;
				}
			}
			const stops = childElements(target).filter((child) => child.localName === 'stop');
			const stop = stops[Math.floor(stops.length / 2)];
			this.degraded.add(gradient ? 'gradientFlatFill' : 'gradientOutlineFlat');
			if (!stop) return null;
			const color = parseColor(stop.getAttribute('stop-color') ?? '#000000');
			return color && { ...color, a: color.a * number(stop.getAttribute('stop-opacity'), 1) };
		}
		if ((target as Element | undefined)?.localName === 'pattern')
			throw new PdfFallback('pattern fill');
		return null;
	}
	walk(element: Element, alpha: number): void {
		if (++this.#nodes > MAX_NODES) throw new PdfFallback('too many drawing nodes');
		const name = element.localName;
		if (SKIPPED.has(name) || element.getAttribute('display') === 'none') return;
		const shadow = element.hasAttribute('data-shadow');
		const glow = element.hasAttribute('data-glow');
		if (glow) {
			this.degraded.add('glowApproximated');
			this.#glow = this.#glowWidth(element);
		}
		if (element.hasAttribute('data-reflection') || element.hasAttribute('mask')) {
			this.degraded.add('reflectionDropped');
			return;
		}
		if (element.hasAttribute('data-bevel')) {
			this.degraded.add('bevelDropped');
			return;
		}
		if (element.hasAttribute('data-group3d-perspective'))
			this.degraded.add('groupPerspectiveAffine');
		if (element.hasAttribute('data-solid3d')) this.degraded.add('solid3dFlat');
		else if (element.hasAttribute('data-perspective')) this.degraded.add('perspectiveParallel');
		if (element.hasAttribute('data-embedded-chart')) throw new PdfFallback('embedded chart');
		const filter = element.hasAttribute('filter') && !glow;
		const soft = filter
			? readShadowFilter(this.#byId(urlId(element.getAttribute('filter'))))
			: undefined;
		const opacity = alpha * number(element.getAttribute('opacity'), 1) * (glow ? 0.5 : 1);
		const transforms = parseTransform(element.getAttribute('transform'));
		const clip = this.#clip(element);
		const wrap = transforms.length > 0 || !!clip;
		if (wrap) {
			this.page.save();
			for (const matrix of transforms) this.page.concat(matrix);
			if (clip) {
				this.page.rect(clip[0], clip[1], clip[2], clip[3]);
				this.page.op('W n');
			}
		}
		// A blurred shadow group is drawn as its soft silhouette; a drop shadow goes beneath.
		const painted = soft
			? paintSvgShadow(this.page, element, soft, opacity, this.fonts, this.familyOf)
			: 'none';
		if (filter && painted === 'none')
			this.degraded.add(shadow ? 'shadowBlurDropped' : 'filterDropped');
		if (painted === 'replaced') {
			if (wrap) this.page.restore();
			return;
		}
		switch (name) {
			case 'g':
			case 'svg':
				for (const child of childElements(element)) this.walk(child, opacity);
				break;
			case 'path':
				this.#path(element, element.getAttribute('d') ?? '', opacity);
				break;
			case 'rect': {
				const [x, y, w, h] = ['x', 'y', 'width', 'height'].map((a) =>
					number(element.getAttribute(a)),
				) as [number, number, number, number];
				if (w > 0 && h > 0)
					this.#path(
						element,
						`M ${x} ${y} L ${x + w} ${y} L ${x + w} ${y + h} L ${x} ${y + h} Z`,
						opacity,
					);
				break;
			}
			case 'text':
				drawSvgText(this.page, element, opacity, this.fonts, this.familyOf);
				break;
			case 'circle':
			case 'ellipse':
			case 'line':
			case 'polyline':
			case 'polygon': {
				const d = basicShapePath(element);
				if (d) this.#path(element, d, opacity);
				break;
			}
			case 'use':
			case 'image':
				this.#image(element, opacity);
				break;
			default:
				throw new PdfFallback(`<${name}> element`);
		}
		if (wrap) this.page.restore();
		if (glow) this.#glow = undefined;
	}
	/** The dilation of a glow's filter, as a stroke width (the radius reaches `size / 2`). */
	#glowWidth(element: Element): number {
		const filter = this.#byId(urlId(element.getAttribute('filter')));
		const radius = number(
			(filter ? childElements(filter) : [])
				.find((child) => child.localName === 'feMorphology')
				?.getAttribute('radius') ?? null,
		);
		return Math.max(0, radius * 2);
	}
	#clip(element: Element): [number, number, number, number] | undefined {
		const id = urlId(element.getAttribute('clip-path'));
		if (!id) return undefined;
		const target = this.#byId(id);
		const shapes = target ? childElements(target) : [];
		const rect = shapes[0];
		if (
			target?.localName !== 'clipPath' ||
			shapes.length !== 1 ||
			rect?.localName !== 'rect' ||
			rect.hasAttribute('transform')
		)
			throw new PdfFallback('clip path');
		return [
			number(rect.getAttribute('x')),
			number(rect.getAttribute('y')),
			number(rect.getAttribute('width')),
			number(rect.getAttribute('height')),
		];
	}
	#path(element: Element, d: string, alpha: number): void {
		const segments = parsePath(d);
		if (!segments) throw new PdfFallback('path data');
		if (!segments.length) return;
		const bounds = pathBounds(segments);
		const painted = this.#paint(element.getAttribute('fill') ?? '#000000', bounds, true);
		const gradient = painted && 'shading' in painted ? painted : undefined;
		const fill = gradient ? { r: 0, g: 0, b: 0, a: gradient.opacity } : (painted as Rgba | null);
		const flat = this.#paint(element.getAttribute('stroke')) as Rgba | null;
		// A glow is a halo: the shape's own colour, stroked wide.
		const stroke = this.#glow !== undefined ? (fill ?? flat) : flat;
		const fillAlpha = fill ? alpha * fill.a * number(element.getAttribute('fill-opacity'), 1) : 0;
		const strokeAlpha = stroke
			? alpha * stroke.a * number(element.getAttribute('stroke-opacity'), 1)
			: 0;
		const width =
			this.#glow !== undefined
				? this.#glow + number(element.getAttribute('stroke-width'), 0)
				: number(element.getAttribute('stroke-width'), 1);
		const page = this.page;
		const doFill = !!fill && fillAlpha > 0,
			doStroke = !!stroke && strokeAlpha > 0 && width > 0;
		const evenOdd = element.getAttribute('fill-rule') === 'evenodd' ? '*' : '';
		if (gradient && doFill) {
			page.save();
			if (fillAlpha < 1) page.alpha(fillAlpha, 1);
			for (const operator of pathOperators(segments)) page.op(operator);
			page.op(`W${evenOdd} n`);
			for (const matrix of gradient.matrices) page.concat(matrix);
			page.paintShading(gradient.shading);
			page.restore();
		}
		if (doFill || doStroke) {
			const flatFill = doFill && !gradient;
			page.save();
			if (fillAlpha < 1 || strokeAlpha < 1) page.alpha(flatFill ? fillAlpha : 1, strokeAlpha);
			if (flatFill) page.fillColor(fill.r, fill.g, fill.b);
			if (doStroke) {
				page.strokeColor(stroke.r, stroke.g, stroke.b);
				const join = this.#glow !== undefined ? 'round' : element.getAttribute('stroke-linejoin');
				const cap = element.getAttribute('stroke-linecap');
				const dash = (element.getAttribute('stroke-dasharray') ?? '')
					.split(/[\s,]+/)
					.map(Number)
					.filter((v) => Number.isFinite(v) && v >= 0);
				page.lineStyle(
					width,
					cap === 'round' ? 1 : cap === 'square' ? 2 : 0,
					join === 'round' ? 1 : join === 'bevel' ? 2 : 0,
					dash.length && dash.some((v) => v > 0) ? dash : [],
				);
			}
			if (flatFill || doStroke) {
				for (const operator of pathOperators(segments)) page.op(operator);
				page.op(flatFill && doStroke ? `B${evenOdd}` : flatFill ? `f${evenOdd}` : 'S');
			}
			page.restore();
		}
		if (stroke && doStroke) this.#markers(element, segments, alpha);
	}
	#markers(element: Element, segments: readonly PathSegment[], alpha: number): void {
		const ends = pathEnds(segments);
		for (const side of ['start', 'end'] as const) {
			const marker = this.#byId(urlId(element.getAttribute(`marker-${side}`)));
			const at = ends[side];
			if (!marker || marker.localName !== 'marker' || !at) continue;
			const [minX, minY, vw, vh] = (marker.getAttribute('viewBox') ?? '0 0 3 3')
				.split(/[\s,]+/)
				.map(Number) as [number, number, number, number];
			void minX;
			void minY;
			const mw = number(marker.getAttribute('markerWidth'), 3),
				mh = number(marker.getAttribute('markerHeight'), 3);
			const scale = Math.min(mw / vw, mh / vh);
			if (!Number.isFinite(scale) || scale <= 0) continue;
			const angle = side === 'start' ? at.angle + Math.PI : at.angle;
			const cos = Math.cos(angle),
				sin = Math.sin(angle);
			this.page.save();
			this.page.concat([cos, sin, -sin, cos, at.x, at.y]);
			this.page.concat([scale, 0, 0, scale, 0, 0]);
			this.page.concat([
				1,
				0,
				0,
				1,
				-number(marker.getAttribute('refX')),
				-number(marker.getAttribute('refY')),
			]);
			for (const child of childElements(marker)) this.walk(child, alpha);
			this.page.restore();
		}
	}
	#image(element: Element, alpha: number): void {
		const href = imageHref(this.root, element);
		const resource = href ? this.images.get(href) : undefined;
		if (!resource) throw new PdfFallback('picture that could not be decoded');
		const w = number(element.getAttribute('width'), 0),
			h = number(element.getAttribute('height'), 0);
		if (!(w > 0 && h > 0)) return;
		const x = number(element.getAttribute('x')),
			y = number(element.getAttribute('y'));
		this.page.save();
		if (alpha < 1) this.page.alpha(alpha);
		this.page.concat([w, 0, 0, -h, x, y + h]);
		this.page.drawImage(resource);
		this.page.restore();
	}
}

/**
 * Convert the static page SVG to a PDF page `widthIn` x `heightIn` inches. Never throws for
 * unsupported content: it returns `fallback` reasons instead.
 */
export function convertSvgPage(
	svg: Element,
	widthIn: number,
	heightIn: number,
	images: ReadonlyMap<string, PdfImageResource>,
	fonts?: PdfFontSet,
	familyOf?: SvgFontFamilyMapper,
): SvgPageConversion {
	const page = new PdfVectorPage(widthIn * 72, heightIn * 72);
	const converter = new Converter(svg, page, images, fonts, familyOf);
	try {
		// SVG user space is y-down inches: map it to PDF points, origin bottom left.
		page.concat([72, 0, 0, -72, 0, page.height]);
		for (const child of childElements(svg)) converter.walk(child, 1);
		return {
			page,
			fallback: [],
			degraded: [...converter.degraded].map((code) => PDF_SVG_NOTES[code]),
		};
	} catch (error) {
		if (error instanceof PdfFallback)
			return { fallback: [`Unsupported content: ${error.message}.`], degraded: [] };
		throw error;
	}
}

/**
 * Draw the children of one `<svg>` root into `page` in the page's current space (SVG user
 * units). Throws `PdfFallback` for content the vector path cannot represent, so a caller can
 * draw that subtree as a picture instead; returns what was simplified, as codes
 * (`PDF_SVG_NOTES` has their English).
 */
export function drawSvgSubtree(
	page: PdfVectorPage,
	svg: Element,
	images: ReadonlyMap<string, PdfImageResource>,
	fonts?: PdfFontSet,
	alpha = 1,
): PdfSvgNoteCode[] {
	const converter = new Converter(svg, page, images, fonts);
	for (const child of childElements(svg)) converter.walk(child, alpha);
	return [...converter.degraded];
}
