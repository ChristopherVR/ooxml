// Pasted ink: the pasted drawing ends up selected, as a pasted floating object is, and a copy of
// floating ink that would land exactly on a drawing already there is moved down and right a
// little, so it is seen as a copy and can be dragged off.
import { Fragment, Slice, type Node as ProseMirrorNode } from 'prosemirror-model';
import { NodeSelection, Plugin } from 'prosemirror-state';

/** How far a copy that would cover its original moves (CSS pixels). */
const PASTE_OFFSET = 12;

const placementOf = (node: ProseMirrorNode): Record<string, unknown> | undefined => {
	try {
		return typeof node.attrs.placement === 'string'
			? (JSON.parse(node.attrs.placement) as Record<string, unknown>)
			: undefined;
	} catch {
		return undefined;
	}
};

/** The positions of the floating drawings already in `doc`, by `x,y,from,fromV`. */
function floatingSpots(doc: ProseMirrorNode): Set<string> {
	const spots = new Set<string>();
	doc.descendants((node) => {
		if (node.type.name !== 'image' || !node.attrs.anchored) return true;
		const placement = placementOf(node);
		if (placement) spots.add(spotOf(placement));
		return false;
	});
	return spots;
}

const spotOf = (placement: Record<string, unknown>) =>
	`${String(placement.offsetXPx)},${String(placement.offsetYPx)},${String(placement.relativeFrom)},${String(placement.relativeFromV)}`;

function offsetInk(fragment: Fragment, taken: Set<string>): Fragment {
	const nodes: ProseMirrorNode[] = [];
	fragment.forEach((node) => {
		if (node.type.name === 'image' && node.attrs.ink && node.attrs.anchored) {
			const placement = placementOf(node);
			if (placement && taken.has(spotOf(placement)) && typeof placement.offsetXPx === 'number') {
				const moved = {
					...placement,
					offsetXPx: placement.offsetXPx + PASTE_OFFSET,
					offsetYPx:
						(typeof placement.offsetYPx === 'number' ? placement.offsetYPx : 0) + PASTE_OFFSET,
				};
				nodes.push(
					node.type.create({ ...node.attrs, placement: JSON.stringify(moved) }, null, node.marks),
				);
				return;
			}
		}
		nodes.push(node.isLeaf ? node : node.copy(offsetInk(node.content, taken)));
	});
	return Fragment.from(nodes);
}

export function inkPastePlugin(): Plugin {
	return new Plugin({
		props: {
			transformPasted(slice, view) {
				let ink = false;
				slice.content.descendants((node) => {
					ink ||= Boolean(node.attrs.ink);
					return !ink;
				});
				if (!ink) return slice;
				return new Slice(
					offsetInk(slice.content, floatingSpots(view.state.doc)),
					slice.openStart,
					slice.openEnd,
				);
			},
		},
		appendTransaction(transactions, _old, state) {
			const paste = transactions.find((tr) => tr.getMeta('uiEvent') === 'paste' && tr.docChanged);
			if (!paste) return null;
			// The ink the paste inserted: the last one between the paste's start and the caret.
			const end = state.selection.from;
			let start = end;
			paste.mapping.maps.forEach((map) =>
				map.forEach((_oldStart, _oldEnd, newStart) => {
					start = Math.min(start, newStart);
				}),
			);
			let last: number | undefined;
			state.doc.nodesBetween(Math.max(0, start), end, (node, pos) => {
				if (node.type.name === 'image' && node.attrs.ink) last = pos;
				return true;
			});
			return last === undefined
				? null
				: state.tr.setSelection(NodeSelection.create(state.doc, last));
		},
	});
}
