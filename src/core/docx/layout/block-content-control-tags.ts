import { cellSourceBlocks } from '../cell-blocks';
import { chainOf } from '../content-controls/chains';
import type { Block, ContentControl, Paragraph } from '../model';

export interface BlockContentControlTags {
	start: ContentControl[];
	end: ContentControl[];
}

/** Row, cell and block controls can each cover several adjacent model nodes. */
export function blockContentControlTags(
	stories: Iterable<readonly Block[]>,
): ReadonlyMap<Paragraph, BlockContentControlTags> {
	const tags = new Map<Paragraph, BlockContentControlTags>();
	const paragraphs: Paragraph[] = [];
	const tagsOf = (paragraph: Paragraph) => {
		let edges = tags.get(paragraph);
		if (!edges) {
			edges = { start: [], end: [] };
			tags.set(paragraph, edges);
		}
		return edges;
	};

	/**
	 * Groups adjacent items that share a tagged control. `controlsOf` lists the controls around an
	 * item, outermost first, so every nesting level gets its own start and end tag.
	 */
	function visitGroups<T>(
		items: readonly T[],
		controlsOf: (item: T, index: number) => readonly ContentControl[],
		visit: (item: T) => void,
	) {
		const active: { control: ContentControl; start: number }[] = [];
		const closeFrom = (keep: number) => {
			while (active.length > keep) {
				const { control, start } = active.pop()!;
				const first = paragraphs[start];
				const last = paragraphs.at(-1);
				if (!first || !last) continue;
				// Inner controls close first (and deeper groups before them), so outer tags enclose them.
				tagsOf(first).start.unshift(control);
				tagsOf(last).end.push(control);
			}
		};
		for (const [index, item] of items.entries()) {
			const chain = controlsOf(item, index).filter((control) => control.appearance === 'tags');
			let keep = 0;
			while (
				keep < active.length &&
				keep < chain.length &&
				active[keep]!.control.id === chain[keep]!.id
			)
				keep++;
			closeFrom(keep);
			for (const control of chain.slice(keep)) active.push({ control, start: paragraphs.length });
			visit(item);
		}
		closeFrom(0);
	}

	function visitBlocks(blocks: readonly Block[]) {
		visitGroups(blocks, chainOf, (block) => {
			if (block.type === 'paragraph') {
				paragraphs.push(block);
				return;
			}
			visitGroups(
				block.rows,
				(_, index) => chainOf(block.rowProperties?.[index]),
				(row) =>
					visitGroups(
						row,
						(cell) => (cell.contentControl ? [cell.contentControl] : []),
						(cell) => visitBlocks(cellSourceBlocks(cell)),
					),
			);
		});
	}

	for (const blocks of stories) {
		visitBlocks(blocks);
		paragraphs.length = 0;
	}
	return tags;
}
