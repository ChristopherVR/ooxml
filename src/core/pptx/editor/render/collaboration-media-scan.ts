/**
 * collaboration-media-scan.ts: which media tokens a collaboration room still refers to.
 *
 * Used before new bytes are added to `pptx:media`, so the entries of pictures nobody uses
 * any more (a replaced deck, a deleted picture) do not count against the room's budget and
 * do not reach the next joiner. The walk is structural (no `yjs` import): maps, arrays and
 * text, whose formatting attributes can name a picture too.
 */

import type { YDocLike } from './collaboration-sync';

const TOKENS = /ooxml-media:([0-9a-f]{32})/gu;

interface Walkable {
	forEach?: (cb: (value: unknown, key: unknown) => void) => void;
	toArray?: () => unknown[];
	toDelta?: () => { insert?: unknown; attributes?: Record<string, unknown> }[];
}

function collect(text: string, into: Set<string>): void {
	if (!text.includes('ooxml-media:')) {
		return;
	}
	for (const match of text.matchAll(TOKENS)) {
		into.add(match[1]!);
	}
}

function walk(value: unknown, into: Set<string>, depth: number): void {
	if (typeof value === 'string') {
		collect(value, into);
		return;
	}
	if (typeof value !== 'object' || value === null || depth > 64) {
		return;
	}
	const node = value as Walkable;
	if (typeof node.toDelta === 'function') {
		for (const op of node.toDelta()) {
			for (const attribute of Object.values(op.attributes ?? {})) {
				walk(attribute, into, depth + 1);
			}
		}
		return;
	}
	if (typeof node.toArray === 'function') {
		for (const item of node.toArray()) {
			walk(item, into, depth + 1);
		}
		return;
	}
	if (typeof node.forEach === 'function') {
		node.forEach((item) => walk(item, into, depth + 1));
	}
}

/** The ids of every media token in the room's slides and assets. */
export function referencedMediaIds(doc: YDocLike): Set<string> {
	const ids = new Set<string>();
	walk(doc.getArray('pptx:slides'), ids, 0);
	walk(doc.getMap('pptx:assets'), ids, 0);
	return ids;
}
