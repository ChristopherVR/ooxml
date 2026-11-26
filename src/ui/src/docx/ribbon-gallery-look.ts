import { PAGE_SIZES } from 'ooxml-core/docx';

/** What a gallery entry looks like: a page thumbnail and a line of detail under the name. */
export interface GalleryLook {
	thumb: SVGSVGElement;
	detail?: string;
}
export type GalleryKind =
	| 'margins'
	| 'size'
	| 'orientation'
	| 'columns'
	| 'verticalAlign'
	| 'borders'
	| 'pageNumber'
	| 'lineSpacing'
	| 'drawingPosition';

const NS = 'http://www.w3.org/2000/svg';
const inches = (twips: number) => `${+(twips / 1440).toFixed(2)}"`;

function shape(name: string, attrs: Record<string, string | number>): SVGElement {
	const el = document.createElementNS(NS, name);
	for (const [key, value] of Object.entries(attrs)) el.setAttribute(key, String(value));
	return el;
}

/** A page outline `w` by `h`, plus whatever `body` draws inside it. */
function page(w: number, h: number, ...body: SVGElement[]): SVGSVGElement {
	const svg = document.createElementNS(NS, 'svg');
	svg.setAttribute('viewBox', `0 0 ${w + 2} ${h + 2}`);
	svg.setAttribute('width', String((w + 2) * 1.4));
	svg.setAttribute('height', String((h + 2) * 1.4));
	svg.setAttribute('aria-hidden', 'true');
	svg.classList.add('gallery-thumb');
	svg.append(shape('rect', { x: 1, y: 1, width: w, height: h, class: 'thumb-page' }), ...body);
	return svg;
}

/** `count` text lines spread over the box `x`, `y`, `w` by `h`. */
const lines = (x: number, y: number, w: number, h: number, count = 5) =>
	Array.from({ length: count }, (_, i) =>
		shape('line', {
			x1: x,
			x2: x + Math.max(w, 1),
			y1: y + (h * (i + 0.5)) / count,
			y2: y + (h * (i + 0.5)) / count,
			class: 'thumb-line',
		}),
	);

/** Word's Margins presets: top and bottom, then left and right, in inches. */
const MARGIN_INCHES: Record<string, [top: number, side: number]> = {
	normal: [1, 1],
	narrow: [0.5, 0.5],
	moderate: [1, 0.75],
	wide: [1, 2],
};

/** A small square with the edges `value` draws in solid ink and the others dotted. */
function borderThumb(value: string): SVGSVGElement {
	const svg = document.createElementNS(NS, 'svg');
	svg.setAttribute('viewBox', '0 0 20 20');
	svg.setAttribute('width', '22');
	svg.setAttribute('height', '22');
	svg.setAttribute('aria-hidden', 'true');
	svg.classList.add('gallery-thumb');
	const on: Record<string, string[]> = {
		bottom: ['b'],
		top: ['t'],
		left: ['l'],
		right: ['r'],
		all: ['t', 'b', 'l', 'r', 'h', 'v'],
		outside: ['t', 'b', 'l', 'r'],
		insideH: ['h'],
		insideV: ['v'],
		horizontal: ['h'],
		none: [],
	};
	const edges: Record<string, [number, number, number, number]> = {
		t: [3, 3, 17, 3],
		b: [3, 17, 17, 17],
		l: [3, 3, 3, 17],
		r: [17, 3, 17, 17],
		h: [3, 10, 17, 10],
		v: [10, 3, 10, 17],
	};
	const active = new Set(on[value] ?? []);
	for (const [edge, [x1, y1, x2, y2]] of Object.entries(edges))
		svg.append(
			shape('line', {
				x1,
				y1,
				x2,
				y2,
				class: active.has(edge) ? 'thumb-edge-on' : 'thumb-edge-off',
			}),
		);
	return svg;
}

/** Three lines spaced like the chosen line spacing, over a small arrow marking the gap. */
function spacingThumb(value: string): SVGSVGElement {
	const twips = Number(/^auto:(\d+)$/.exec(value)?.[1] ?? 240);
	const gap = Math.max(3, Math.min(9, (twips / 240) * 3.5));
	const svg = document.createElementNS(NS, 'svg');
	svg.setAttribute('viewBox', '0 0 24 24');
	svg.setAttribute('width', '24');
	svg.setAttribute('height', '24');
	svg.setAttribute('aria-hidden', 'true');
	svg.classList.add('gallery-thumb');
	const top = 12 - gap;
	for (const y of [top, 12, 12 + gap])
		svg.append(shape('line', { x1: 8, x2: 21, y1: y, y2: y, class: 'thumb-line thumb-line-dark' }));
	svg.append(shape('path', { d: 'M4 4v16M2 6l2-2 2 2M2 18l2 2 2-2', class: 'thumb-arrow' }));
	return svg;
}

/** The thumbnail and detail for `value` of the gallery `kind`; unknown values get a plain page. */
export function galleryLook(
	kind: GalleryKind,
	value: string,
	say: (text: string) => string = (text) => text,
): GalleryLook {
	if (kind === 'drawingPosition') {
		const [v, h] = value.split('-');
		const x = h === 'left' ? 5 : h === 'right' ? 21 : 13;
		const y = v === 'top' ? 5 : v === 'bottom' ? 28 : 16;
		return {
			thumb: page(
				32,
				40,
				...lines(5, 5, 24, 32, 9),
				shape('rect', { x, y, width: 8, height: 8, class: 'thumb-page thumb-edge-on' }),
			),
			detail: value === 'inline' ? say('In Line with Text') : say('With Text Wrapping'),
		};
	}
	if (kind === 'margins') {
		const [top, side] = MARGIN_INCHES[value] ?? [1, 1];
		const k = 4;
		return {
			thumb: page(30, 40, ...lines(1 + side * k, 1 + top * k, 30 - side * 2 * k, 40 - top * 2 * k)),
			detail: `${say('Top')}: ${top}"  ${say('Bottom')}: ${top}"  ${say('Left')}: ${side}"  ${say('Right')}: ${side}"`,
		};
	}
	if (kind === 'size') {
		const [w, h] = PAGE_SIZES[value] ?? [12240, 15840];
		const scale = 44 / 24480;
		const width = Math.round(w * scale);
		const height = Math.round(h * scale);
		return {
			thumb: page(width, height, ...lines(4, 5, width - 6, height - 9, 4)),
			detail: `${inches(w)} x ${inches(h)}`,
		};
	}
	if (kind === 'orientation') {
		return {
			thumb:
				value === 'landscape'
					? page(44, 32, ...lines(5, 5, 34, 22, 4))
					: page(32, 44, ...lines(5, 5, 22, 34, 6)),
		};
	}
	if (kind === 'columns') {
		if (value === 'left' || value === 'right') {
			const first = value === 'left' ? 6 : 13;
			return {
				thumb: page(
					30,
					40,
					...lines(4, 5, first, 30, 6),
					...lines(4 + first + 3, 5, 19 - first, 30, 6),
				),
			};
		}
		const count = Number(value) || 1;
		const gap = 3;
		const width = (30 - 8 - gap * (count - 1)) / count;
		const columns = Array.from({ length: count }, (_, i) =>
			lines(4 + i * (width + gap), 5, width, 30, 6),
		);
		return { thumb: page(30, 40, ...columns.flat()) };
	}
	if (kind === 'pageNumber') {
		const [position, align, format] = value.split(':');
		const number = shape('text', {
			x: align === 'left' ? 5 : align === 'right' ? 27 : 16,
			y: position === 'top' ? 6 : 39,
			'font-size': 3,
			'text-anchor': align === 'left' ? 'start' : align === 'right' ? 'end' : 'middle',
			fill: 'currentColor',
		});
		number.textContent = format === 'pageOfTotal' ? '1 / 2' : '1';
		return { thumb: page(30, 40, ...lines(5, 10, 22, 22, 6), number) };
	}
	if (kind === 'borders') return { thumb: borderThumb(value) };
	if (kind === 'lineSpacing') return { thumb: spacingThumb(value) };
	const top = { top: 5, center: 15, bottom: 25, both: 5 }[value as 'top'] ?? 5;
	const both = value === 'both';
	return { thumb: page(30, 40, ...lines(5, top, 20, both ? 30 : 10, both ? 7 : 3)) };
}
