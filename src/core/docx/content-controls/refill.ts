// Filling an already filled document again: its structure comes back from the template it was
// filled from. In each story, every block that holds a content control is the template's (so an
// optional block a previous fill removed is back where it was, and repeated copies go, to be made
// again from the new items), and every block without one is the document's: the text the user
// typed outside the controls stays, between the controlled blocks it followed. The template's
// blocks get ids of their own and are written from the model.
import { mapStoryBlocks } from '../cell-blocks';
import { mapBlockLists } from '../mail-merge/stories';
import type { Block, DocumentModel } from '../model';
import { chainOf, groupAt } from './chains';
import { forEachControl } from './controls';

interface Node {
	blocks: Block[];
	/** The id of the control that holds it, or of the first control inside it. */
	key?: string;
}

function nodesOf(blocks: readonly Block[]): Node[] {
	return groupAt(blocks, 0, chainOf).flatMap((group): Node[] => {
		if (group.control) return [{ blocks: group.items, key: group.control.id }];
		return group.items.map((block) => {
			let key: string | undefined;
			forEachControl([block], (control) => (key ??= control.id));
			return key === undefined ? { blocks: [block] } : { blocks: [block], key };
		});
	});
}

/** The story lists of a model (not cells or text boxes), in the order they are mapped. */
function storyLists(model: DocumentModel): Block[][] {
	const lists: Block[][] = [];
	mapBlockLists(model, (blocks, list) => {
		if (!list.cell && list.story !== 'textBox') lists.push(blocks);
		return blocks;
	});
	return lists;
}

function merge(
	document: readonly Block[],
	template: readonly Block[],
	rekey: (id: string) => string,
): Block[] {
	const own = nodesOf(document);
	const theirs = nodesOf(template);
	const output: Block[] = [];
	let cursor = 0;
	const keep = (until: number) => {
		for (; cursor < until; cursor++) {
			const node = own[cursor]!;
			if (node.key === undefined) output.push(...node.blocks);
		}
	};
	const copy = (node: Node) =>
		mapStoryBlocks(node.blocks, (paragraph) => ({ ...paragraph, id: rekey(paragraph.id) }), rekey);
	// A block the document lost (an optional block removed) comes back before the next block both
	// have, after the text typed before that one; at the end, before the trailing text.
	let pending: Node[] = [];
	for (const node of theirs) {
		if (node.key === undefined) continue;
		const at = own.findIndex((item, index) => index >= cursor && item.key === node.key);
		if (at < 0) {
			pending.push(node);
			continue;
		}
		keep(at);
		cursor = at + 1;
		output.push(...pending.flatMap(copy), ...copy(node));
		pending = [];
	}
	output.push(...pending.flatMap(copy));
	keep(own.length);
	return output.length ? output : [...document];
}

/** `document` with the structure of `template` (see above). */
export function reconcile(document: DocumentModel, template: DocumentModel): DocumentModel {
	const lists = storyLists(template);
	const ids: string[] = [];
	mapBlockLists(document, (blocks) => {
		for (const block of blocks) ids.push(block.id);
		return blocks;
	});
	let suffix = '-tpl';
	while (ids.some((id) => id.includes(suffix))) suffix += 'l';
	const rekey = (id: string) => `${id}${suffix}`;
	let index = 0;
	const reconciled = mapBlockLists(document, (blocks, list) => {
		if (list.cell || list.story === 'textBox') return blocks;
		const theirs = lists[index++];
		return theirs ? merge(blocks, theirs, rekey) : blocks;
	});
	// The document's lists include the template's (and the ones earlier HTML values added).
	const numbering = document.numberingCatalog ?? template.numberingCatalog;
	return numbering ? { ...reconciled, numberingCatalog: numbering } : reconciled;
}
