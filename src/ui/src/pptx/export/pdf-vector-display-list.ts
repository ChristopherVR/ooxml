/**
 * The display list a mounted slide stage is read into for vector PDF export
 * (`pdf-vector-collect.ts`) and painted from (`pdf-vector-paint.ts`). Every geometry is in
 * CSS px of the element's own untransformed layout, relative to the stage's top-left corner;
 * `matrix` maps that space to the stage's painted space (the element's CSS transforms).
 * DOM-free apart from the `svg` item, which carries a detached SVG tree.
 */
import type { PdfImageResource, PdfShading } from 'ooxml-core/pdf/write';

/** A 2-D affine matrix `[a, b, c, d, e, f]`, as in CSS `matrix()`. */
export type Matrix = readonly [number, number, number, number, number, number];

export const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];

/** `outer` applied after `inner`. */
export function multiply(outer: Matrix, inner: Matrix): Matrix {
	const [a, b, c, d, e, f] = outer;
	const [A, B, C, D, E, F] = inner;
	return [
		a * A + c * B,
		b * A + d * B,
		a * C + c * D,
		b * C + d * D,
		a * E + c * F + e,
		b * E + d * F + f,
	];
}

export interface Rgba {
	r: number;
	g: number;
	b: number;
	a: number;
}

/** A clip region: SVG path data in the space `matrix` maps to the stage. */
export interface PdfClip {
	matrix: Matrix;
	d: string;
	evenOdd?: boolean;
}

interface ItemBase {
	matrix: Matrix;
	/** Clips that apply, outermost first. */
	clips: readonly PdfClip[];
	/** Accumulated opacity. */
	alpha: number;
}

export interface PathStroke {
	color: Rgba;
	width: number;
	dash: readonly number[];
	/** 0 butt, 1 round, 2 square. */
	cap: number;
	/** 0 miter, 1 round, 2 bevel. */
	join: number;
}

/** A filled and/or stroked path (box backgrounds, borders, CSS clip shapes). */
export interface PathItem extends ItemBase {
	kind: 'path';
	d: string;
	evenOdd?: boolean;
	fill?: Rgba;
	/** A gradient fill: painted inside the path after `matrices` are applied. */
	shading?: { shading: PdfShading; matrices: readonly Matrix[]; opacity: number };
	stroke?: PathStroke;
}

/** A picture drawn into the rectangle `x, y, width, height`. */
export interface ImageItem extends ItemBase {
	kind: 'image';
	x: number;
	y: number;
	width: number;
	height: number;
	image: PdfImageResource;
}

/** One line fragment of text, fitted to its measured width. */
export interface TextItem extends ItemBase {
	kind: 'text';
	text: string;
	x: number;
	baseline: number;
	/** The measured advance of the fragment on the slide. */
	width: number;
	size: number;
	/** The CSS `font-family` list. */
	family: string;
	bold: boolean;
	italic: boolean;
	color: Rgba;
	underline?: boolean;
	strike?: boolean;
	/** Selectable and searchable but not drawn: a picture of it is underneath. */
	invisible?: boolean;
}

/** An inline SVG drawn by the shared SVG converter, in its own user units. */
export interface SvgItem extends ItemBase {
	kind: 'svg';
	svg: Element;
	images: ReadonlyMap<string, PdfImageResource>;
}

export type DisplayItem = PathItem | ImageItem | TextItem | SvgItem;
