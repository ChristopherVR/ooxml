import { Fragment, Slice, type Node, type ResolvedPos } from 'prosemirror-model';
import { Plugin, Selection, SelectionRange, type SelectionBookmark } from 'prosemirror-state';
import type { Mappable } from 'prosemirror-transform';
import { Decoration, DecorationSet } from 'prosemirror-view';

function cellBookmark(ranges: { from: number; to: number }[], anchor: number): SelectionBookmark {
	return {
		map(mapping) {
			const mapped = ranges.flatMap((range) => {
				const from = mapping.mapResult(range.from, 1);
				const to = mapping.mapResult(range.to, -1);
				return from.deletedAcross || to.deletedAcross || from.pos > to.pos
					? []
					: [{ from: from.pos, to: to.pos }];
			});
			return cellBookmark(mapped, mapping.map(anchor));
		},
		resolve(doc) {
			return ranges.length
				? fromRanges(
						ranges.map(({ from, to }) => new SelectionRange(doc.resolve(from), doc.resolve(to))),
						doc,
					)
				: Selection.near(doc.resolve(anchor));
		},
	};
}

/** A column is a disjoint selection. Keep each range limited to one cell. */
export class TableCellSelection extends Selection {
	override visible = false;

	constructor(ranges: SelectionRange[], anchor: ResolvedPos, head: ResolvedPos) {
		super(anchor, head, ranges);
	}

	override eq(other: Selection): boolean {
		return (
			other instanceof TableCellSelection &&
			this.ranges.length === other.ranges.length &&
			this.ranges.every(
				(range, index) =>
					range.$from.pos === other.ranges[index]?.$from.pos &&
					range.$to.pos === other.ranges[index]?.$to.pos,
			)
		);
	}

	override map(doc: Node, mapping: Mappable): Selection {
		return this.getBookmark().map(mapping).resolve(doc);
	}

	override getBookmark(): SelectionBookmark {
		return cellBookmark(
			this.ranges.map(({ $from, $to }) => ({ from: $from.pos, to: $to.pos })),
			this.anchor,
		);
	}

	override content(): Slice {
		const rows = new Map<number, { node: Node; cells: Node[] }>();
		let table: Node | undefined;
		for (const { $from } of this.ranges) {
			let depth = $from.depth;
			while (depth > 2 && $from.node(depth).type.name !== 'tableCell') depth--;
			if ($from.node(depth).type.name !== 'tableCell') return super.content();
			table ??= $from.node(depth - 2);
			if ($from.node(depth - 2) !== table) return super.content();
			const pos = $from.before(depth - 1);
			let row = rows.get(pos);
			if (!row) rows.set(pos, (row = { node: $from.node(depth - 1), cells: [] }));
			const cell = $from.node(depth);
			row.cells.push(
				cell.type.create(
					{ ...cell.attrs, colspan: 1, rowspan: 1, continueKeys: null },
					cell.content,
				),
			);
		}
		if (!table) return Slice.empty;
		const content = [...rows.values()].map(({ node, cells }) =>
			node.copy(Fragment.fromArray(cells)),
		);
		return new Slice(
			Fragment.from(table.type.create({ ...table.attrs, grid: null }, content)),
			0,
			0,
		);
	}

	override toJSON() {
		return {
			type: 'dve-table-cells',
			ranges: this.ranges.map(({ $from, $to }) => [$from.pos, $to.pos]),
		};
	}

	static override fromJSON(doc: Node, json: unknown): Selection {
		const values = json && typeof json === 'object' && 'ranges' in json ? json.ranges : null;
		if (!Array.isArray(values) || !values.length || values.length > 10_000)
			throw new RangeError('Invalid table cell selection');
		const ranges = values.map((pair: unknown) => {
			if (
				!Array.isArray(pair) ||
				pair.length !== 2 ||
				!pair.every(Number.isSafeInteger) ||
				pair[0] < 0 ||
				pair[0] > pair[1] ||
				pair[1] > doc.content.size
			)
				throw new RangeError('Invalid table cell range');
			return new SelectionRange(doc.resolve(pair[0]), doc.resolve(pair[1]));
		});
		return fromRanges(ranges, doc);
	}
}
Selection.jsonID('dve-table-cells', TableCellSelection);

export function fromRanges(ranges: SelectionRange[], doc: Node): Selection {
	const first = ranges[0];
	return first ? new TableCellSelection(ranges, first.$from, first.$to) : Selection.atStart(doc);
}

export function selectCellNodes(doc: Node, table: Node, tablePos: number, cells: Set<Node>) {
	const ranges: SelectionRange[] = [];
	table.forEach((row, rowOffset) => {
		row.forEach((cell, cellOffset) => {
			if (!cells.has(cell)) return;
			const pos = tablePos + rowOffset + cellOffset + 2;
			ranges.push(
				new SelectionRange(
					Selection.near(doc.resolve(pos + 1), 1).$from,
					Selection.near(doc.resolve(pos + cell.nodeSize - 1), -1).$to,
				),
			);
		});
	});
	return ranges.length ? fromRanges(ranges, doc) : Selection.near(doc.resolve(tablePos + 2));
}

export function tableCellSelectionPlugin(): Plugin {
	return new Plugin({
		props: {
			decorations({ selection, doc }) {
				if (!(selection instanceof TableCellSelection)) return null;
				const decorations = selection.ranges.flatMap(({ $from }) => {
					for (let depth = $from.depth; depth > 0; depth--)
						if ($from.node(depth).type.name === 'tableCell')
							return [
								Decoration.node($from.before(depth), $from.after(depth), {
									class: 'dve-selected-table-cell',
								}),
							];
					return [];
				});
				return DecorationSet.create(doc, decorations);
			},
		},
	});
}
