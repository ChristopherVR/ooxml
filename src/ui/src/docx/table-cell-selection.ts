import type { Node, ResolvedPos } from 'prosemirror-model';
import { Plugin, Selection, SelectionRange } from 'prosemirror-state';
import type { Mappable } from 'prosemirror-transform';
import { Decoration, DecorationSet } from 'prosemirror-view';

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
		const ranges = this.ranges.flatMap((range) => {
			const from = mapping.mapResult(range.$from.pos, 1);
			const to = mapping.mapResult(range.$to.pos, -1);
			return from.deletedAcross || to.deletedAcross || from.pos > to.pos
				? []
				: [new SelectionRange(doc.resolve(from.pos), doc.resolve(to.pos))];
		});
		return ranges.length
			? fromRanges(ranges, doc)
			: Selection.near(doc.resolve(mapping.map(this.anchor)));
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
