import type { Mark, ResolvedPos } from 'prosemirror-model';
import { TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import type { Decoration, DecorationSet } from 'prosemirror-view';
import { controlOfMark, isWordYjsRemoteTransaction } from 'ooxml-core/docx/ui';
import { inlineTagsAt, type TagEdge } from './content-control-tag-widgets';

const controlMark = (mark: Mark) => mark.type.name === 'contentControl';
const markAt = ($pos: ResolvedPos, edge: TagEdge) =>
	(edge === 'start' ? $pos.nodeAfter : $pos.nodeBefore)?.marks.find(controlMark);
const opposite = (edge: TagEdge): TagEdge => (edge === 'start' ? 'end' : 'start');

export function inlineCaretSide(
	widgets: DecorationSet,
	selection: EditorState['selection'],
	prior?: { tr: Transaction; state: EditorState; inside: TagEdge | undefined },
	clicked?: TagEdge,
): TagEdge | undefined {
	if (!selection.empty || !(selection instanceof TextSelection)) return undefined;
	const $pos = selection.$head;
	const tags = inlineTagsAt(widgets, $pos.pos);
	const has = (edge: TagEdge) => tags.some(({ spec }) => spec.edge === edge);
	if (!tags.length) return undefined;
	if (clicked && has(clicked)) return clicked;

	if (prior?.state.selection instanceof TextSelection && prior.state.selection.empty) {
		const { $head } = prior.state.selection;
		const { inside } = prior;
		// A Yjs replacement restores the selection independently of the step mapping.
		if (isWordYjsRemoteTransaction(prior.tr)) {
			if (!inside || !has(inside)) return undefined;
			const id = controlOfMark(markAt($head, inside))?.id;
			return id !== undefined && controlOfMark(markAt($pos, inside))?.id === id
				? inside
				: undefined;
		}
		const previous = prior.tr.mapping.map($head.pos, -1);
		// Text typed at the inside end moves the caret with the end tag: it stays inside.
		const followed = prior.tr.mapping.map($head.pos, 1);
		if ((previous === $pos.pos || followed === $pos.pos) && inside && has(inside)) return inside;
		if (previous !== $pos.pos && $pos.sameParent($pos.doc.resolve(previous))) {
			const edge = previous > $pos.pos ? 'start' : 'end';
			const approach = markAt($head, opposite(edge));
			const target = markAt($pos, edge);
			if (target && has(edge) && approach?.eq(target)) return edge;
		}
	}
	return undefined;
}

/** Match inherited marks to the visible side of a tagged inline boundary when typing there. */
export function inlineCaretMarks(
	widgets: DecorationSet,
	inside: TagEdge | undefined,
	state: EditorState,
): Transaction | undefined {
	const { selection, storedMarks } = state;
	if (!selection.empty || !(selection instanceof TextSelection)) return undefined;
	const $pos = selection.$head;
	const edges = inlineTagsAt(widgets, $pos.pos);
	if (!edges.length) return undefined;
	const current = storedMarks ?? $pos.marks();
	const mark = current.find(controlMark);
	const belongsToEdge = ({ spec }: Decoration) => {
		const edge: unknown = spec.edge;
		const edgeMark = edge === 'start' || edge === 'end' ? markAt($pos, edge) : undefined;
		return edgeMark && mark?.eq(edgeMark);
	};
	const desired = inside ? markAt($pos, inside) : edges.some(belongsToEdge) ? undefined : mark;
	if (mark === desired || (mark && desired && mark.eq(desired))) return undefined;
	const keep = mark ? mark.removeFromSet(current) : current;
	return state.tr.setStoredMarks(desired ? desired.addToSet(keep) : keep);
}
