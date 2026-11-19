// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Puts rewritten paragraphs, tables, rows, cells and grid columns where the ones they replace were.
// xmldom re-indexes every child of a parent on each insertion or removal that is not an append, so
// moving every block of a long document is quadratic: only the nodes that must move are moved.
import type { XmlElement } from './xml';

/**
 * Replaces `oldNodes` slot by slot with `nextNodes`: the i-th next node takes the place of the
 * i-th old node, in whichever parent that old node sits (a content control's, say). Old nodes
 * without a counterpart are removed, and the remaining next nodes are inserted in `parent` before
 * `anchor` (at the end without one). A node listed twice in `nextNodes` ends up in its last slot.
 */
export function replaceSlots(
	parent: XmlElement,
	oldNodes: readonly XmlElement[],
	nextNodes: readonly XmlElement[],
	anchor: XmlElement | null = null,
): void {
	if (oldNodes.length === nextNodes.length && oldNodes.every((node, i) => node === nextNodes[i]))
		return;
	const lastSlot = new Map<Node, number>();
	nextNodes.forEach((node, index) => lastSlot.set(node, index));
	const placedAt = (index: number): XmlElement | undefined => {
		const node = nextNodes[index];
		return node && lastSlot.get(node) === index ? node : undefined;
	};
	const slotOf = new Map<Node, number>();
	const owners = new Set<XmlElement>([parent]);
	oldNodes.forEach((node, index) => {
		const owner = node.parentNode as XmlElement | null;
		if (!owner) throw new Error('DOCX block slot is not in the document.');
		slotOf.set(node, index);
		owners.add(owner);
	});
	const extras: XmlElement[] = [];
	for (let index = oldNodes.length; index < nextNodes.length; index++) {
		const node = placedAt(index);
		if (node) extras.push(node);
	}
	if (extras.length && anchor && anchor.parentNode !== parent)
		throw new Error('DOCX block anchor is not in its container.');
	// Every target is worked out before any node moves, from the children as they are now.
	const targets = [...owners].map((owner) => {
		const target: Node[] = [];
		for (let child = owner.firstChild; child; child = child.nextSibling) {
			if (owner === parent && child === anchor) target.push(...extras);
			const slot = slotOf.get(child);
			if (slot !== undefined) {
				const next = placedAt(slot);
				if (next) target.push(next);
			} else if (!lastSlot.has(child)) target.push(child);
		}
		if (owner === parent && !anchor) target.push(...extras);
		return { owner, target };
	});
	for (const { owner, target } of targets) reorderChildren(owner, target);
}

/** Gives `owner` exactly the children `target`, moving the fewest nodes. */
function reorderChildren(owner: XmlElement, target: readonly Node[]): void {
	const position = new Map<Node, number>();
	let count = 0;
	for (let child = owner.firstChild; child; child = child.nextSibling) position.set(child, count++);
	if (target.length === count && target.every((node, index) => position.get(node) === index))
		return;
	const wanted = new Set(target);
	for (const child of [...position.keys()]) if (!wanted.has(child)) owner.removeChild(child);
	const stays = longestIncreasing(target.map((node) => position.get(node) ?? -1));
	let next: Node | null = null;
	for (let index = target.length - 1; index >= 0; index--) {
		const node = target[index]!;
		if (!stays.has(index)) owner.insertBefore(node, next);
		next = node;
	}
}

/** Indices of a longest strictly increasing run of `values`; negative values never join it. */
function longestIncreasing(values: readonly number[]): Set<number> {
	// tails[k]: index of the smallest value ending an increasing run of length k + 1.
	const tails: number[] = [];
	const previous = new Array<number>(values.length).fill(-1);
	values.forEach((value, index) => {
		if (value < 0) return;
		let low = 0;
		let high = tails.length;
		while (low < high) {
			const middle = (low + high) >> 1;
			if (values[tails[middle]!]! < value) low = middle + 1;
			else high = middle;
		}
		previous[index] = low > 0 ? tails[low - 1]! : -1;
		tails[low] = index;
	});
	const stays = new Set<number>();
	for (let index = tails.at(-1) ?? -1; index >= 0; index = previous[index]!) stays.add(index);
	return stays;
}
