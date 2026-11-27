import { Fragment, type Node } from 'prosemirror-model';
import { TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import { createClientId } from '../../collab/identity';
import type { Revision } from '../model';
import { createCollaborationIdGenerator } from './collaboration-identity';
import { outlinePeer, type OutlineEntry } from './outline';
import { formattingRevision } from './review-formatting';
import { inlineTextRevision } from './review-inline-revisions';
import { trackChangesPluginKey } from './track-changes-mode';

/**
 * Why an outline branch move cannot be recorded as a tracked change: the branch holds more than
 * plain text paragraphs (tables, pictures, fields, breaks), it already carries tracked changes,
 * or no paragraph follows the place it leaves or the place it lands (Word keeps a moved paragraph
 * mark only before another paragraph, and accepting or rejecting merges into that paragraph).
 */
export type TrackedOutlineMoveRefusal = 'content' | 'revisions' | 'boundary';

export interface TrackedOutlineMoveOptions {
	author: string;
	date?: string;
	/** Revision and paragraph ids; collaboration passes its room generator. */
	nextId?: (kind: 'revision' | 'paragraph') => string;
}

function branchRefusal(nodes: readonly Node[]): TrackedOutlineMoveRefusal | undefined {
	for (const node of nodes) {
		if (node.type.name !== 'paragraph') return 'content';
		if (node.attrs.markRevision || node.attrs.formatRevision) return 'revisions';
		let refusal: TrackedOutlineMoveRefusal | undefined;
		node.forEach((child) => {
			if (refusal) return;
			if (!child.isText) refusal = 'content';
			else if (inlineTextRevision(child) || formattingRevision(child)) refusal = 'revisions';
		});
		if (refusal) return refusal;
	}
	return undefined;
}

function placement(
	state: EditorState,
	entries: readonly OutlineEntry[],
	entry: OutlineEntry,
	direction: -1 | 1,
) {
	const peer = outlinePeer(entries, entry, direction);
	if (!peer) return undefined;
	const size = state.doc.content.size;
	const sourceEnd = entries[entry.end]?.pos ?? size;
	const insertAt = direction === 1 ? (entries[peer.end]?.pos ?? size) : peer.pos;
	const nodes: Node[] = [];
	state.doc.nodesBetween(entry.pos, sourceEnd, (node, pos) => {
		if (pos >= entry.pos) nodes.push(node);
		return false;
	});
	return { peer, sourceEnd, insertAt, nodes };
}

/** Undefined when moving `entry` one peer up or down can be recorded as a Word move. */
export function trackedOutlineMoveRefusal(
	state: EditorState,
	entries: readonly OutlineEntry[],
	entry: OutlineEntry,
	direction: -1 | 1,
): TrackedOutlineMoveRefusal | undefined {
	const place = placement(state, entries, entry, direction);
	if (!place) return 'boundary';
	const refusal = branchRefusal(place.nodes);
	if (refusal) return refusal;
	const isParagraph = (pos: number) => state.doc.nodeAt(pos)?.type.name === 'paragraph';
	const afterSource = place.sourceEnd < state.doc.content.size && isParagraph(place.sourceEnd);
	const afterTarget = place.insertAt < state.doc.content.size && isParagraph(place.insertAt);
	return afterSource && afterTarget ? undefined : 'boundary';
}

/**
 * Moves a whole outline branch as Word records it with Track Changes on: the branch stays where it
 * was as moved-from text and paragraph marks, and a copy with fresh paragraph ids is inserted at
 * the destination as moved-to content, both sides linked by one move name. With move tracking off
 * (`trackMoves` false) the two sides are an ordinary deletion and insertion.
 */
export function trackedMoveOutline(
	state: EditorState,
	entries: readonly OutlineEntry[],
	entry: OutlineEntry,
	direction: -1 | 1,
	options: TrackedOutlineMoveOptions,
): Transaction | undefined {
	if (trackedOutlineMoveRefusal(state, entries, entry, direction)) return undefined;
	const place = placement(state, entries, entry, direction);
	if (!place) return undefined;
	const { from, to, anchor, head } = state.selection;
	if (from < entry.pos || to >= place.sourceEnd) return undefined;
	const { schema } = state;
	const insertion = schema.marks.insertion;
	const deletion = schema.marks.deletion;
	if (!insertion || !deletion) return undefined;
	const generator = createCollaborationIdGenerator(createClientId());
	const nextId = options.nextId ?? generator;
	const date = options.date ?? new Date(Date.now()).toISOString();
	const trackMoves = state.doc.attrs.trackMoves !== false;
	const name = `move${createClientId().replaceAll('-', '')}`;
	const revision = (kind: 'from' | 'to'): Revision => ({
		kind: trackMoves
			? kind === 'from'
				? 'moveFrom'
				: 'moveTo'
			: kind === 'from'
				? 'delete'
				: 'insert',
		author: options.author,
		date,
		dateUtc: date,
		id: nextId('revision'),
		...(trackMoves ? { move: { name } } : {}),
	});
	const markAttrs = (kind: 'from' | 'to') => {
		const { id, author } = revision(kind);
		return { author, date, dateUtc: date, id, move: trackMoves ? JSON.stringify({ name }) : null };
	};
	const used = new Set<string>();
	state.doc.descendants((node) => {
		if (node.attrs.id) used.add(String(node.attrs.id));
	});
	const freshId = () => {
		let id: string;
		do id = nextId('paragraph');
		while (used.has(id));
		used.add(id);
		return id;
	};

	const moved = insertion.create(markAttrs('to'));
	const copies = place.nodes.map((node) => {
		const content: Node[] = [];
		node.forEach((child) => content.push(child.mark(moved.addToSet(child.marks))));
		return node.type.create(
			{ ...node.attrs, id: freshId(), markRevision: revision('to') },
			Fragment.fromArray(content),
			node.marks,
		);
	});
	const tr = state.tr;
	const removed = deletion.create(markAttrs('from'));
	let pos = entry.pos;
	for (const node of place.nodes) {
		if (node.content.size) tr.addMark(pos + 1, pos + node.nodeSize - 1, removed);
		tr.setNodeMarkup(pos, undefined, { ...node.attrs, markRevision: revision('from') });
		pos += node.nodeSize;
	}
	tr.insert(place.insertAt, Fragment.fromArray(copies));
	const shift = place.insertAt - entry.pos;
	tr.setSelection(
		TextSelection.between(tr.doc.resolve(anchor + shift), tr.doc.resolve(head + shift)),
	);
	return tr.setMeta(trackChangesPluginKey, { tracked: true });
}
