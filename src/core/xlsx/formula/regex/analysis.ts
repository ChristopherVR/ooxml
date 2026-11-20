import type { CharMatcher } from './charset';
import { Assertion, type Node } from './parse';

/** Whether `node` can match the empty string. */
export function nullable(node: Node): boolean {
	switch (node.kind) {
		case 'char':
			return false;
		case 'seq':
			return node.items.every(nullable);
		case 'alt':
			return node.items.some(nullable);
		case 'group':
			return nullable(node.body);
		case 'repeat':
			return node.min === 0 || nullable(node.body);
		default:
			return true;
	}
}

interface First {
	readonly set: CharMatcher[] | null;
	readonly empty: boolean;
}

/** The atoms a match of `node` can start with (null: unknown), and whether it can match empty. */
export function firstAtoms(node: Node): First {
	switch (node.kind) {
		case 'char':
			return { set: [node.matcher], empty: false };
		case 'group':
			return firstAtoms(node.body);
		case 'repeat': {
			if (node.max === 0) return { set: [], empty: true };
			const inner = firstAtoms(node.body);
			return { set: inner.set, empty: node.min === 0 || inner.empty };
		}
		case 'seq': {
			const set: CharMatcher[] = [];
			for (const item of node.items) {
				const next = firstAtoms(item);
				if (!next.set) return next;
				set.push(...next.set);
				if (!next.empty) return { set, empty: false };
			}
			return { set, empty: true };
		}
		case 'alt': {
			const set: CharMatcher[] = [];
			let empty = false;
			for (const item of node.items) {
				const next = firstAtoms(item);
				if (!next.set) return next;
				set.push(...next.set);
				empty ||= next.empty;
			}
			return { set, empty };
		}
		case 'backref':
			return { set: null, empty: true };
		default:
			return { set: [], empty: true };
	}
}

/** Whether every match of `node` starts at position 0. */
export function anchored(node: Node): boolean {
	if (node.kind === 'assert') return node.what === Assertion.Start;
	if (node.kind === 'seq') return node.items[0] !== undefined && anchored(node.items[0]);
	if (node.kind === 'group') return anchored(node.body);
	return node.kind === 'alt' && node.items.every(anchored);
}
