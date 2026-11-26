import type { Mark, ResolvedPos } from 'prosemirror-model';
import { TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import type { Decoration, DecorationSet } from 'prosemirror-view';
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
		const previous = prior.tr.mapping.map($head.pos, -1);
		if (previous === $pos.pos && prior.inside && has(prior.inside)) return prior.inside;
		if (previous !== $pos.pos && $pos.sameParent($pos.doc.resolve(previous))) {
			const edge = previous > $pos.pos ? 'start' : 'end';
			const approach = markAt($head, opposite(edge));
			const target = markAt($pos, edge);
			if (target && has(edge) && approach?.eq(target)) return edge;
		}
	}
	if ($pos.parentOffset === 0 && has('start')) return 'start';
	if ($pos.parentOffset === $pos.parent.content.size && has('end')) return 'end';
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
