/**
 * A shape placed on a page, the options it is placed with, and the options of a connector
 * between two of them.
 */
import type { ConnectorLabel, Direction, Point, Port } from './connector';
import { type Size, along } from './geometry';
import type { ShapeDef } from './shape-def';
import type { Rich } from './text';

export interface Link {
	/** A page of this drawing. */
	page?: string;
	/** An external address. */
	address?: string;
	description: string;
}

export interface NodeOptions {
	x: number;
	y: number;
	w?: number;
	h?: number;
	text?: Rich;
	name?: string;
	/** Local cells written as given. */
	cells?: readonly string[];
	/** Local cells for the sub-shapes of a group master's instance, by master shape ID. */
	childCells?: Readonly<Record<number, readonly string[]>>;
	/** Values for the Shape Data rows the master (or the drawn shape) declares. */
	data?: Readonly<Record<string, string | number>>;
	layers?: readonly string[];
	link?: Link;
}

export interface ConnectOptions {
	via?: readonly Point[];
	/** Keep the corners of the route square (lines that share a run, as in a tree). */
	square?: boolean;
	label?: string | ConnectorLabel;
	layers?: readonly string[];
	cells?: readonly string[];
}

/** A placed 2-D shape and the connection points connectors glue to. */
export class Node {
	constructor(
		readonly id: number,
		readonly x: number,
		readonly y: number,
		readonly size: Size,
		private readonly def: ShapeDef,
	) {}

	get left(): number {
		return this.x - this.size.w / 2;
	}
	get right(): number {
		return this.x + this.size.w / 2;
	}
	get top(): number {
		return this.y + this.size.h / 2;
	}
	get bottom(): number {
		return this.y - this.size.h / 2;
	}

	/** The connection point with the given row index. */
	port(index: number): Port {
		const point = this.def.connections?.[index];
		if (!point) throw new Error(`Shape ${this.id} has no connection point ${index}.`);
		const fx = along(point[0], this.size.w) / this.size.w;
		const fy = along(point[1], this.size.h) / this.size.h;
		const distances: [Direction, number][] = [
			['w', fx],
			['e', 1 - fx],
			['s', fy],
			['n', 1 - fy],
		];
		distances.sort((a, b) => a[1] - b[1]);
		return {
			shapeId: this.id,
			index,
			x: this.left + fx * this.size.w,
			y: this.bottom + fy * this.size.h,
			dir: distances[0]?.[0] ?? 'e',
		};
	}
}
