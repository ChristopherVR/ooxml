/**
 * Builds the shapes of one page: instances of the drawing's masters, shapes drawn on the page,
 * glued connectors, and the edits the core library applies afterwards (shape data, hyperlinks,
 * layers).
 */
import type { VisioEdit } from '../../../../src/core/visio/edit-commands';
import type { VisioShapeDataType } from '../../../../src/core/visio/edit-shape-data-commands';
import { xmlEscape } from '../../lib/package';
import { type Port, connectorShape, route } from './connector';
import { rectangle } from './geometry';
import { type PlacedMaster, type Stencil, MASTER_ROOT_ID, instanceIdCount } from './masters';
import {
	type PropDef,
	type ShapeDef,
	PROP,
	drawnShape,
	instanceShape,
	propertySection,
} from './shape';
import type { StyleBook } from './styles';
import { type Rich, hasField, hyperlinkSection, textBody } from './text';
import { attrs, cell, mergeCells } from './xml';

import { type ConnectOptions, type Link, type NodeOptions, Node } from './node';

export { type ConnectOptions, type Link, type NodeOptions, Node } from './node';

const LIBRARY_TYPES: Record<number, VisioShapeDataType> = {
	[PROP.string]: 'string',
	[PROP.list]: 'fixed-list',
	[PROP.number]: 'number',
	[PROP.boolean]: 'boolean',
	[PROP.date]: 'date',
};

export class PageBuilder {
	private shapes: string[] = [];
	private connects: string[] = [];
	/** Library edits to run once the shapes are in the page part. */
	edits: VisioEdit[] = [];
	/** Top-level shape IDs per layer set (names joined by `|`), for the library to assign. */
	readonly layerSets = new Map<string, number[]>();
	/** The page's layers in the order of their rows; a layer's index is its ID. */
	readonly layerNames: string[] = [];
	/** Masters this page has instances of, for the page's relationships. */
	readonly usedMasters = new Set<number>();
	/**
	 * Text with fields, held back per shape ID. The library refuses to link data to a page on
	 * which a formula already reads the linked rows, so such text is first written plain and
	 * the fields are added once the links exist (`Drawing.applyFields`).
	 */
	readonly deferred = new Map<number, Rich>();
	deferFields = false;
	private next = 1;

	constructor(
		readonly pageId: string,
		readonly name: string,
		readonly styles: StyleBook,
		private readonly stencil: Stencil,
	) {}

	/** Takes `count` shape IDs for shapes the library creates; returns them. */
	reserve(count = 1): string[] {
		const ids: string[] = [];
		for (let index = 0; index < count; index++) ids.push(String(this.next++));
		return ids;
	}

	/** Puts top-level shapes (also ones the library created) on layers. */
	assign(layers: readonly string[], ids: readonly (number | string)[]): void {
		for (const id of ids) this.track(Number(id), { layers });
	}

	private layerIds(layers: readonly string[]): number[] {
		return layers.map((name) => {
			if (!this.layerNames.includes(name)) this.layerNames.push(name);
			return this.layerNames.indexOf(name);
		});
	}

	/**
	 * The LayerMember cell of a group and its sub-shapes. The library does not assign groups to
	 * layers, so an instance of a group master carries its membership from the start.
	 */
	private groupLayers(
		layers: readonly string[] | undefined,
		subShapes: number,
	): { cells: string[]; childCells: Record<number, string[]> } {
		if (!layers?.length) return { cells: [], childCells: {} };
		const member = [cell('LayerMember', this.layerIds(layers).join(';'))];
		const childCells: Record<number, string[]> = {};
		for (let index = 1; index <= subShapes; index++) childCells[MASTER_ROOT_ID + index] = member;
		return { cells: member, childCells };
	}

	private track(id: number, options: { layers?: readonly string[] }): void {
		if (!options.layers?.length) return;
		this.layerIds(options.layers);
		const key = options.layers.join('|');
		const members = this.layerSets.get(key) ?? [];
		members.push(id);
		this.layerSets.set(key, members);
	}

	/**
	 * Queues the library edits for a shape's data and hyperlink, and returns the rows the
	 * library has no type for (and an instance's hyperlink, which it refuses) written by hand.
	 */
	private metadata(
		id: number,
		props: readonly PropDef[],
		options: NodeOptions,
		instance: boolean,
	): string {
		const shapeId = String(id);
		const byHand: PropDef[] = [];
		for (const [name, value] of Object.entries(options.data ?? {})) {
			const prop = props.find((entry) => entry.name === name);
			if (!prop) throw new Error(`Shape ${id} has no Shape Data row ${name}.`);
			const type = LIBRARY_TYPES[prop.type];
			if (!type) {
				byHand.push({ ...prop, value });
				continue;
			}
			this.edits.push({
				type: 'set-shape-data',
				pageId: this.pageId,
				shapeId,
				row: name,
				data: {
					label: prop.label,
					prompt: prop.prompt ?? '',
					type,
					format: prop.format ?? '',
					value: String(value),
				},
			});
		}
		const rows = instance
			? propertySection(byHand.map((prop) => ({ ...prop, valueOnly: true })))
			: propertySection([
					...props.filter((prop) => options.data?.[prop.name] === undefined),
					...byHand,
				]);
		if (!options.link) return rows;
		const hyperlink = {
			address: options.link.address ?? '',
			subAddress: options.link.page ?? '',
			description: options.link.description,
		};
		// The library refuses hyperlinks on master instances, so those rows are written here.
		if (instance) return rows + hyperlinkSection(hyperlink);
		this.edits.push({ type: 'set-shape-hyperlink', pageId: this.pageId, shapeId, hyperlink });
		return rows;
	}

	/** The text to write now: plain when its fields are held back. */
	private textNow(id: number, rich: Rich | undefined): Rich | undefined {
		if (!this.deferFields || !hasField(rich)) return rich;
		this.deferred.set(id, rich);
		return rich.map((run) => run.text).join('');
	}

	/** An instance of a master, centred on (x, y). */
	drop(masterName: string, options: NodeOptions): Node {
		const master = this.stencil.get(masterName);
		if (master.custom) throw new Error(`${masterName} is a connector master; use connect().`);
		const size = { w: options.w ?? master.size.w, h: options.h ?? master.size.h };
		const id = this.next;
		let masterShape = MASTER_ROOT_ID;
		const name = options.name ?? this.instanceName(master, id);
		const rows = this.metadata(id, master.shape.props ?? [], options, true);
		const subShapes = instanceIdCount(master) - 1;
		const grouped = this.groupLayers(subShapes ? options.layers : undefined, subShapes);
		const childCells: Record<number, readonly string[]> = { ...grouped.childCells };
		for (const [key, cells] of Object.entries(options.childCells ?? {}))
			childCells[Number(key)] = mergeCells(childCells[Number(key)], cells);
		const shape = instanceShape(
			master.shape,
			master.size,
			{ x: options.x, y: options.y, size },
			() => this.next++,
			() => masterShape++,
			{
				head: attrs({ NameU: name, Name: name, Master: master.id }),
				styles: this.styles,
				cells: mergeCells(grouped.cells, options.cells),
				childCells,
				body: rows + textBody(this.textNow(id, options.text)),
			},
		);
		this.shapes.push(shape);
		this.usedMasters.add(master.id);
		if (!subShapes) this.track(id, options);
		return new Node(id, options.x, options.y, size, master.shape);
	}

	/** A shape drawn on the page, with its own geometry. */
	draw(def: ShapeDef, options: NodeOptions): Node {
		const size = { w: options.w ?? 1, h: options.h ?? 1 };
		const id = this.next;
		const named = options.name ? attrs({ NameU: options.name, Name: options.name }) : '';
		// Rows with values go through the library; the shape itself declares the others.
		const { props, ...rest } = def;
		const rows = this.metadata(id, props ?? [], options, false);
		const shape = drawnShape(
			{ ...rest, cells: [...(def.cells ?? []), ...(options.cells ?? [])] },
			{ x: options.x, y: options.y, size },
			this.styles,
			() => this.next++,
			{ head: named, body: rows + textBody(this.textNow(id, options.text)), onPage: true },
		);
		this.shapes.push(shape);
		this.track(id, options);
		return new Node(id, options.x, options.y, size, def);
	}

	/** Text with no outline or fill, in a box whose left edge is at `left`. */
	note(
		style: string,
		text: Rich,
		box: { left: number; y: number; w: number; h?: number; name?: string; link?: Link },
	): Node {
		return this.draw(
			{ style, paths: [rectangle(0, 0, 1, 1, { noFill: true, noLine: true })] },
			{
				x: box.left + box.w / 2,
				y: box.y,
				w: box.w,
				h: box.h ?? 0.25,
				text,
				...(box.name ? { name: box.name } : {}),
				...(box.link ? { link: box.link } : {}),
			},
		);
	}

	/** A paragraph of body text that starts at `top` and wraps to `w`. */
	paragraph(text: string, box: { left: number; top: number; w: number; h?: number }): Node {
		const h = box.h ?? 0.6;
		return this.note('Body', text, { left: box.left, y: box.top - h / 2, w: box.w, h });
	}

	/** A filled box with no text: rules, bands and swatches. */
	block(style: string, box: { left: number; bottom: number; w: number; h: number }): Node {
		return this.draw(
			{ style, paths: [rectangle()] },
			{ x: box.left + box.w / 2, y: box.bottom + box.h / 2, w: box.w, h: box.h },
		);
	}

	private instanceName(master: PlacedMaster, id: number): string {
		master.uses++;
		return master.uses === 1 ? master.name : `${master.name}.${id}`;
	}

	/** A connector glued to two connection points, routed at right angles. */
	connect(masterName: string, from: Port, to: Port, options: ConnectOptions = {}): void {
		const master = this.stencil.get(masterName);
		if (instanceIdCount(master) !== 1 || !master.custom)
			throw new Error(`${masterName} is not a connector master.`);
		const id = this.next++;
		const label = typeof options.label === 'string' ? { text: options.label } : options.label;
		const { shape, connects } = connectorShape({
			id,
			master,
			name: this.instanceName(master, id),
			from,
			to,
			points: route(from, to, options.via),
			...(label ? { label: { ...label, text: xmlEscape(label.text) } } : {}),
			cells: mergeCells(this.styles.themeCells(master.shape.style), options.cells),
			...(options.square ? { square: true } : {}),
		});
		this.shapes.push(shape);
		this.connects.push(connects);
		this.usedMasters.add(master.id);
		this.track(id, options);
	}

	/** The shapes and Connect records built since the last call. */
	flush(): { shapes: string; connects: string } {
		const result = { shapes: this.shapes.join(''), connects: this.connects.join('') };
		this.shapes = [];
		this.connects = [];
		return result;
	}

	takeEdits(): VisioEdit[] {
		const edits = this.edits;
		this.edits = [];
		return edits;
	}
}
