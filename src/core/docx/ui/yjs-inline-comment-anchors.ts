import * as Y from 'yjs';
import { Plugin, type EditorState } from 'prosemirror-state';
import { Fragment, type Node } from 'prosemirror-model';
import {
	absolutePositionToRelativePosition,
	relativePositionToAbsolutePosition,
	ySyncPluginKey,
} from 'y-prosemirror';
import { commentIdsFromNode } from './comment-anchors';
import { commentSelectionRange } from './comment-selection';

type Mapping = Parameters<typeof absolutePositionToRelativePosition>[2];
interface AnchorRange {
	from: number[];
	to: number[];
}
const supported = (node: Node) =>
	node.isInline && !node.isText && Boolean(node.type.spec.attrs?.format);

/** Each thread owns its element ranges. Independent thread keys merge without replacing a group. */
export class WordYjsInlineCommentAnchors {
	readonly ranges: Y.Map<AnchorRange[]>;
	constructor(
		private readonly fragment: Y.XmlFragment,
		private readonly deleted: Y.Map<boolean>,
	) {
		this.ranges = fragment.doc!.getMap('docx:comments:inline-anchors');
	}

	private capture(doc: Node, from: number, to: number, mapping: Mapping): AnchorRange[] {
		const ranges: AnchorRange[] = [];
		const relative = (pos: number, assoc: number): number[] => {
			const position = absolutePositionToRelativePosition(pos, this.fragment, mapping);
			const absolute = Y.createAbsolutePositionFromRelativePosition(position, this.fragment.doc!);
			if (!absolute) throw new Error('Cannot resolve an inline comment anchor.');
			return Array.from(
				Y.encodeRelativePosition(
					Y.createRelativePositionFromTypeIndex(absolute.type, absolute.index, assoc),
				),
			);
		};
		doc.nodesBetween(from, to, (node, pos) => {
			if (!supported(node)) return;
			ranges.push({
				from: relative(pos, 0),
				to: relative(pos + node.nodeSize, -1),
			});
		});
		return ranges;
	}

	seed(doc: Node, mapping: Mapping): void {
		const byId = new Map<string, AnchorRange[]>();
		doc.descendants((node, pos) => {
			if (!supported(node)) return;
			const ranges = this.capture(doc, pos, pos + node.nodeSize, mapping);
			for (const id of commentIdsFromNode(node)) byId.set(id, [...(byId.get(id) ?? []), ...ranges]);
		});
		for (const [id, ranges] of byId) this.ranges.set(id, ranges);
	}

	add(state: EditorState, id: string): boolean {
		const binding = ySyncPluginKey.getState(state)?.binding;
		if (!binding || binding.type !== this.fragment) return false;
		const selection = commentSelectionRange(state.doc, state.selection.from, state.selection.to);
		const ranges = this.capture(state.doc, selection.from, selection.to, binding.mapping);
		if (!ranges.length) return false;
		this.ranges.set(id, ranges);
		return true;
	}

	project(doc: Node, mapping: Mapping): Node {
		const markType = doc.type.schema.marks.inlineCommentAnchors;
		if (!markType) return doc;
		const idsAt = new Map<number, string[]>();
		for (const [id, ranges] of this.ranges) {
			if (this.deleted.get(id)) continue;
			for (const range of ranges) {
				const from = relativePositionToAbsolutePosition(
					this.fragment.doc!,
					this.fragment,
					Y.decodeRelativePosition(Uint8Array.from(range.from)),
					mapping,
				);
				const to = relativePositionToAbsolutePosition(
					this.fragment.doc!,
					this.fragment,
					Y.decodeRelativePosition(Uint8Array.from(range.to)),
					mapping,
				);
				if (
					from === null ||
					to === null ||
					to - from !== 1 ||
					!doc.nodeAt(from) ||
					!supported(doc.nodeAt(from)!)
				)
					continue;
				idsAt.set(from, [...new Set([...(idsAt.get(from) ?? []), id])].sort());
			}
		}
		const visit = (node: Node, pos: number): Node => {
			if (supported(node)) {
				const marks = node.marks.filter(
					(mark) => mark.type !== markType && mark.type.name !== 'comment',
				);
				const ids = idsAt.get(pos) ?? [];
				return node.mark(
					ids.length || commentIdsFromNode(node.mark(marks)).length
						? [...marks, markType.create({ ids })]
						: marks,
				);
			}
			if (node.isLeaf) return node;
			const children: Node[] = [];
			node.forEach((child, offset) =>
				children.push(visit(child, pos + offset + (node.type === doc.type ? 0 : 1))),
			);
			return node.copy(Fragment.fromArray(children));
		};
		return visit(doc, 0);
	}

	plugin(): Plugin {
		return new Plugin({
			appendTransaction: (_transactions, _old, state) => {
				const binding = ySyncPluginKey.getState(state)?.binding;
				if (!binding) return;
				const doc = this.project(state.doc, binding.mapping);
				if (doc.eq(state.doc)) return;
				const tr = state.tr;
				state.doc.descendants((node, pos) => {
					if (!supported(node)) return;
					const projected = doc.nodeAt(pos)!;
					for (const mark of node.marks)
						if (!projected.marks.some((next) => next.eq(mark))) tr.removeNodeMark(pos, mark);
					for (const mark of projected.marks)
						if (!node.marks.some((previous) => previous.eq(mark))) tr.addNodeMark(pos, mark);
				});
				return tr.setMeta('addToHistory', false).setMeta('dve-remote', true);
			},
			view: (view) => {
				const refresh = () => {
					const binding = ySyncPluginKey.getState(view.state)?.binding;
					// Projection is read-only. A map observer may run before the fragment observer;
					// writing its older text snapshot back would erase incoming text comment marks.
					binding?.mux(() =>
						view.dispatch(view.state.tr.setMeta('word-inline-comment-anchors', true)),
					);
				};
				this.ranges.observe(refresh);
				this.deleted.observe(refresh);
				return {
					destroy: () => {
						this.ranges.unobserve(refresh);
						this.deleted.unobserve(refresh);
					},
				};
			},
		});
	}
}
