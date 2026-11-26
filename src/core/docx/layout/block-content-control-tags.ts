import { cellSourceBlocks } from '../cell-blocks';
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

	function visitGroups<T>(
		items: readonly T[],
		controlOf: (item: T, index: number) => ContentControl | undefined,
		visit: (item: T) => void,
	) {
		let active: ContentControl | undefined;
		let start = paragraphs.length;
		const close = () => {
			const first = paragraphs[start];
			const last = paragraphs.at(-1);
			if (!active || !first || !last) return;
			// The inner control was recorded first. The outer tags enclose it.
			tagsOf(first).start.unshift(active);
			tagsOf(last).end.push(active);
		};
		for (const [index, item] of items.entries()) {
			const control = controlOf(item, index);
			const tagged = control?.appearance === 'tags' ? control : undefined;
			if (tagged?.id !== active?.id) {
				close();
				active = tagged;
				start = paragraphs.length;
			}
			visit(item);
		}
		close();
	}

	function visitBlocks(blocks: readonly Block[]) {
		visitGroups(
			blocks,
			(block) => block.contentControl,
			(block) => {
				if (block.type === 'paragraph') {
					paragraphs.push(block);
					return;
				}
				visitGroups(
					block.rows,
					(_, index) => block.rowProperties?.[index]?.contentControl,
					(row) =>
						visitGroups(
							row,
							(cell) => cell.contentControl,
							(cell) => visitBlocks(cellSourceBlocks(cell)),
						),
				);
			},
		);
	}

	for (const blocks of stories) {
		visitBlocks(blocks);
		paragraphs.length = 0;
	}
	return tags;
}
