/**
 * What a 2-D shape is described by: its style, outline, text block, connection points and
 * Shape Data. `shape.ts` writes a description as a master, a drawn shape or an instance.
 */
import type { Co, Path } from './geometry';

/** Shape Data types (the Type cell of a Property row). */
export const PROP = { string: 0, list: 1, number: 2, boolean: 3, date: 5, currency: 7 } as const;

export interface PropDef {
	/** Row name, read as `Prop.<name>`. */
	name: string;
	label: string;
	type: (typeof PROP)[keyof typeof PROP];
	/** A format picture, or the `;` separated choices of a list. */
	format?: string;
	prompt?: string;
	value?: string | number;
	/** An instance overriding the value of a row its master declares. */
	valueOnly?: boolean;
}

export interface UserDef {
	name: string;
	value: string | number;
	unit?: string;
	formula?: string;
}

/** Where the text block sits: its centre and size as fractions (plus inches) of the shape. */
export interface TextBlock {
	x: Co;
	y: Co;
	w: Co;
	h: Co;
}

export interface ShapeDef {
	/** One style name for line, fill and text, or one per kind. */
	style: string | { line: string; fill: string; text: string };
	paths: readonly Path[];
	textBlock?: TextBlock;
	/** Connection points as fractions of the shape; rows are glued to by index. */
	connections?: readonly (readonly [Co, Co])[];
	props?: readonly PropDef[];
	user?: readonly UserDef[];
	/** Extra ShapeSheet cells written as given. */
	cells?: readonly string[];
	/** Sub-shapes; their presence makes the shape a group. */
	children?: readonly ChildDef[];
}

export interface ChildDef extends ShapeDef {
	/** Left, bottom, width and height inside the parent. */
	box: { x: Co; y: Co; w: Co; h: Co };
}

/** The twelve points every box-like master offers: mid-sides first, then the quarter points. */
export const BOX_POINTS: readonly (readonly [number, number])[] = [
	[0.5, 1],
	[1, 0.5],
	[0.5, 0],
	[0, 0.5],
	[0.25, 1],
	[0.75, 1],
	[1, 0.75],
	[1, 0.25],
	[0.75, 0],
	[0.25, 0],
	[0, 0.25],
	[0, 0.75],
];
export const SIDE = { top: 0, right: 1, bottom: 2, left: 3 } as const;
